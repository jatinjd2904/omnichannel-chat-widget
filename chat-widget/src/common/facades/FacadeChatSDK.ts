import { BroadcastEvent, LogLevel, TelemetryEvent } from "../telemetry/TelemetryConstants";
import { ChatAdapter, ChatSDKMessage, GetAgentAvailabilityResponse, GetLiveChatTranscriptResponse, GetPersistentChatHistoryResponse, GetVoiceVideoCallingResponse, IFileInfo, IRawMessage, MaskingRules, OmnichannelChatSDK, VoiceVideoCallingOptionalParams } from "@microsoft/omnichannel-chat-sdk";
import { IFacadeChatSDKInput, PingResponse } from "./types/IFacadeChatSDKInput";
import { getAuthClientFunction, handleAuthentication } from "../../components/livechatwidget/common/authHelper";

import { BroadcastService } from "@microsoft/omnichannel-chat-components";
import ChatAdapterOptionalParams from "@microsoft/omnichannel-chat-sdk/lib/core/messaging/ChatAdapterOptionalParams";
import ChatConfig from "@microsoft/omnichannel-chat-sdk/lib/core/ChatConfig";
import ChatReconnectContext from "@microsoft/omnichannel-chat-sdk/lib/core/ChatReconnectContext";
import ChatReconnectOptionalParams from "@microsoft/omnichannel-chat-sdk/lib/core/ChatReconnectOptionalParams";
import ChatTranscriptBody from "@microsoft/omnichannel-chat-sdk/lib/core/ChatTranscriptBody";
import EmailLiveChatTranscriptOptionaParams from "@microsoft/omnichannel-chat-sdk/lib/core/EmailLiveChatTranscriptOptionalParams";
import EndChatOptionalParams from "@microsoft/omnichannel-chat-sdk/lib/core/EndChatOptionalParams";
import FileMetadata from "@microsoft/omnichannel-amsclient/lib/FileMetadata";
import GetAgentAvailabilityOptionalParams from "@microsoft/omnichannel-chat-sdk/lib/core/GetAgentAvailabilityOptionalParams";
import GetChatTokenOptionalParams from "@microsoft/omnichannel-chat-sdk/lib/core/GetChatTokenOptionalParams";
import GetConversationDetailsOptionalParams from "@microsoft/omnichannel-chat-sdk/lib/core/GetConversationDetailsOptionalParams";
import GetLiveChatConfigOptionalParams from "@microsoft/omnichannel-chat-sdk/lib/core/GetLiveChatConfigOptionalParams";
import GetLiveChatTranscriptOptionalParams from "@microsoft/omnichannel-chat-sdk/lib/core/GetLiveChatTranscriptOptionalParams";
import GetPersistentChatHistoryOptionalParams from "@microsoft/omnichannel-chat-sdk/lib/core/GetPersistentChatHistoryOptionalParams";
import IChatToken from "@microsoft/omnichannel-chat-sdk/lib/external/IC3Adapter/IChatToken";
import IFileMetadata from "@microsoft/omnichannel-ic3core/lib/model/IFileMetadata";
import IMessage from "@microsoft/omnichannel-ic3core/lib/model/IMessage";
import IRawThread from "@microsoft/omnichannel-ic3core/lib/interfaces/IRawThread";
import InitializeOptionalParams from "@microsoft/omnichannel-chat-sdk/lib/core/InitializeOptionalParams";
import LiveWorkItemDetails from "@microsoft/omnichannel-chat-sdk/lib/core/LiveWorkItemDetails";
import OmnichannelMessage from "@microsoft/omnichannel-chat-sdk/lib/core/messaging/OmnichannelMessage";
import OnNewMessageOptionalParams from "@microsoft/omnichannel-chat-sdk/lib/core/messaging/OnNewMessageOptionalParams";
import { ParticipantsRemovedEvent } from "@azure/communication-signaling";
import PostChatContext from "@microsoft/omnichannel-chat-sdk/lib/core/PostChatContext";
import StartChatOptionalParams from "@microsoft/omnichannel-chat-sdk/lib/core/StartChatOptionalParams";
import { TelemetryHelper } from "../telemetry/TelemetryHelper";
import { isNullOrEmptyString } from "../utils";

export class FacadeChatSDK {
    private chatSDK: OmnichannelChatSDK;
    private chatConfig: ChatConfig;
    private token: string | "" | null = "";
    private expiration = 0;
    private isAuthenticated: boolean;
    private getAuthToken?: (authClientFunction?: string) => Promise<string | null>;
    private sdkMocked: boolean;
    private disableReauthentication: boolean;

    // Flag set by tokenRing() when mid-auth is enabled and no token is available (user not logged in).
    // NOT reset in CASE 1 - stays true so CASE 1 re-triggers on every startChat to set deferInitialAuth.
    private pendingMidAuthUnauthenticatedState = false;
    
    /**
     * Mid-auth enabled check based on chatConfig.
     * Mid-auth flag lives under LiveWSAndLiveChatEngJoin.msdyn_authenticatedsigninoptional.
     */
    private isMidAuthEnabled(): boolean {
        const value = (this.chatConfig as ChatConfig)?.LiveWSAndLiveChatEngJoin?.msdyn_authenticatedsigninoptional;
        return value?.toString?.().toLowerCase?.() === "true";
    }

    public isSDKMocked(): boolean {
        return this.sdkMocked;
    }

    public getChatSDK(): OmnichannelChatSDK {
        return this.chatSDK;
    }

    public destroy() {
        console.info("[LCW][FacadeChatSDK][destroy] Clearing authentication state");
        this.token = null;
        this.expiration = 0;
        // Reset mid-auth state only when mid-auth is enabled
        if (this.isMidAuthEnabled()) {
            this.pendingMidAuthUnauthenticatedState = false;
            this.isAuthenticated = true;
            // Reset deferInitialAuth so next chat doesn't inherit unauthenticated mode
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            (this.chatSDK as any).deferInitialAuth = false;
        }
    }

    public isTokenSet() {
        console.info("[LCW][FacadeChatSDK][isTokenSet] invoked");
        return !isNullOrEmptyString(this.token);
    }

    constructor(input: IFacadeChatSDKInput, disableReauthentication: boolean) {
        console.info("[LCW][FacadeChatSDK][constructor]", {
            hasGetAuthToken: !!input.getAuthToken,
            isAuthenticated: input.isAuthenticated,
            isSDKMocked: input.isSDKMocked,
            disableReauthentication: disableReauthentication
        });
        this.chatSDK = input.chatSDK;
        this.chatConfig = input.chatConfig;
        this.getAuthToken = input.getAuthToken;
        this.isAuthenticated = input.isAuthenticated;
        this.sdkMocked = input.isSDKMocked;
        this.disableReauthentication = disableReauthentication;
    }

    //set default expiration to zero, for undefined or missed exp in jwt
    private convertExpiration(expiration = 0): number {
        console.info("[LCW][FacadeChatSDK][convertExpiration] invoked with expiration:", expiration);
        // Converting expiration to seconds, if contains decimals or is identified as milliseconds
        if (expiration.toString().length === 13) {
            return Math.floor(expiration / 1000);
        }
        // If the epoch value is already in seconds, return it as is
        return expiration;
    }

    private isTokenExpired(): boolean {
        // if expiration is 0, token is not going to be validated ( this is to cover the case of token with no expiration)
        if (this.expiration === 0) {
            return false;
        }

        // obtain current time in seconds
        const now = Math.floor(Date.now() / 1000);

        // compare expiration time with current time
        if (now > this.expiration) {
            console.error("Token is expired", now, this.expiration, now > this.expiration);
            return true;
        }

        return false;
    }

    private enforceBase64Encoding(payload: string): string {
        //base64url when present, switches the "-" and "_" characters with "+" and "/"
        const base64Payload = payload.replace(/-/g, "+").replace(/_/g, "/");
        // since base64 encoding requires padding, we need to add padding to the payload
        return base64Payload.padEnd(base64Payload.length + (4 - base64Payload.length % 4) % 4, "=");
    }

    private extractExpFromToken(token: string): number {

        const tokenParts = token.split(".");
        const last3digits = token.slice(-3);

        // token must have 3 parts as JWT format
        if (tokenParts.length !== 3) {
            TelemetryHelper.logFacadeChatSDKEventToAllTelemetry(LogLevel.ERROR, {
                Event: TelemetryEvent.NewTokenValidationFailed,
                Description: "Invalid token format",
                ExceptionDetails: { message: "Invalid token format, must be in JWT format", token: last3digits }
            });
            throw new Error("Authentication Setup Error: Invalid token format, must be in JWT format");
        }

        try {
            const payload = this.enforceBase64Encoding(tokenParts[1]);
            // decode payload
            const decodedPayload = atob(payload);
            const jsonPayload = JSON.parse(decodedPayload);
            // check if exp is present in payload
            if (jsonPayload) {
                if (jsonPayload.exp) {
                    return jsonPayload.exp;
                }
                return 0;
            }
            TelemetryHelper.logFacadeChatSDKEventToAllTelemetry(LogLevel.ERROR, {
                Event: TelemetryEvent.NewTokenValidationFailed,
                Description: "Invalid token payload",
                ExceptionDetails: { message: "Token payload is not valid JSON", token: last3digits }
            });

            throw new Error("Authentication Setup Error: Invalid token payload, payload is not valid JSON");

        } catch (e) {
            console.error("Authentication Setup Error: Failed to decode token", e);
            TelemetryHelper.logFacadeChatSDKEventToAllTelemetry(LogLevel.ERROR, {
                Event: TelemetryEvent.NewTokenValidationFailed,
                Description: "Failed to decode token",
                ExceptionDetails: { message: "Failed to decode token", token: last3digits }
            });
            throw new Error("Authentication Setup Error: Failed to decode authentication token");
        }
    }

    private async setToken(token: string): Promise<void> {
        console.info("[LCW][FacadeChatSDK][setToken] invoked");
        // token must be not null, and must be new
        if (!isNullOrEmptyString(token) && token !== this.token) {
            console.info("[LCW][FacadeChatSDK][setToken] setting new token");
            const last3digits = token.slice(-3);
            const instant = Math.floor(Date.now() / 1000);
            this.token = token;
            // calculate expiration time
            this.expiration = this.convertExpiration(this.extractExpFromToken(token) || 0);
            // this is a control , in case the getAuthToken function returns same token
            if (this.expiration > 0 && (this.expiration < instant)) {
                TelemetryHelper.logFacadeChatSDKEventToAllTelemetry(LogLevel.ERROR, {
                    Event: TelemetryEvent.NewTokenValidationFailed,
                    Description: "New token is already expired",
                    ExceptionDetails: {
                        "Instant": instant,
                        "Expiration": this.expiration,
                        "Token": last3digits,
                    }
                });
                throw new Error("Authentication Setup Error: New authentication token is already expired");
            }
        }
    }

    private async corroborateTokenIsSet(chatSDK: OmnichannelChatSDK): Promise<void> {
        console.info("[LCW][FacadeChatSDK][corroborateTokenIsSet]", {
            isAuthenticated: this.isAuthenticated,
            hasChatSDKConfigGetAuthToken: !!chatSDK?.chatSDKConfig?.getAuthToken
        });

        // if getAuthToken is not set, it's because handleAuthentication hasnt being called
        // so we need to call it 
        if (this.isAuthenticated && chatSDK?.chatSDKConfig?.getAuthToken === undefined) {
            console.info("[LCW][FacadeChatSDK][corroborateTokenIsSet] calling handleAuthentication");
            handleAuthentication(this.chatSDK, this.chatConfig, this.getAuthToken);
        }
    }
    private async tokenRing(): Promise<PingResponse> {

        // Use console logging for local debugging (telemetry can be delayed/filtered).
        console.info("[LCW][FacadeChatSDK][tokenRing] START", {
            disableReauthentication: this.disableReauthentication,
            sdkMocked: this.sdkMocked,
            isAuthenticated: this.isAuthenticated,
            isTokenSet: this.isTokenSet(),
            isMidAuthEnabled: this.isMidAuthEnabled()
        });

        if (this.disableReauthentication === true) {
            console.info("[LCW][FacadeChatSDK][tokenRing] BRANCH: disableReauthentication=true");
            // Since we are not validating the token anymore, we at least need to check if the token is set
            // no need to validate anything other that the token is set
            await this.corroborateTokenIsSet(this.chatSDK);
            // facade feature is disabled, so we are bypassing the re authentication and let it fail.
            return { result: true, message: "Facade is disabled" };
        }

        // this is needed for storybooks, specifically for reconnect pane which requires authentication bypass
        if (this.sdkMocked === true) {
            console.info("[LCW][FacadeChatSDK][tokenRing] BRANCH: sdkMocked=true");
            return { result: true, message: "Authentication not needed" };
        }

        // If isAuthenticated is false, authentication is not required for this chat
        // This covers: unauthenticated chats, mid-auth before user authenticates, etc.
        if (!this.isAuthenticated) {
            console.info("[LCW][FacadeChatSDK][tokenRing] BRANCH: isAuthenticated=false - authentication not required for this chat");
            return { result: true, message: "Authentication not needed" };
        }

        if (this.isTokenSet() && !this.isTokenExpired()) {
            console.info("[LCW][FacadeChatSDK][tokenRing] BRANCH: token is set and valid");
            return { result: true, message: "Token is valid" };
        }

        // If we reach here, we need to get a token via getAuthToken
        // MID-AUTH: getAuthToken receives { isMidAuthEnabled: true } flag for customer implementations to handle
        // Customer implementations can check portal state and return null for logged-out users
        // For Custom Portals: handleAuthentication will call getAuthToken and handle null/empty/HTML responses
        console.info("[LCW][FacadeChatSDK][tokenRing] BRANCH: auth required - need to get token", {
            authClientFunction: getAuthClientFunction(this.chatConfig),
            hasPropGetAuthToken: this.getAuthToken !== undefined,
            hasSdkGetAuthToken: this.chatSDK?.chatSDKConfig?.getAuthToken !== undefined,
        });

        TelemetryHelper.logFacadeChatSDKEventToAllTelemetry(LogLevel.INFO, {
            Event: TelemetryEvent.NewTokenValidationStarted,
            Description: "Token validation started."
        });

        if (this.getAuthToken === undefined && this.chatSDK.chatSDKConfig?.getAuthToken === undefined) {
            console.info("[LCW][FacadeChatSDK][tokenRing] ERROR: GetAuthToken function is not present", {
                getAuthTokenFromProps: this.getAuthToken,
                getAuthTokenFromSDK: this.chatSDK.chatSDKConfig?.getAuthToken,
                authClientFunction: getAuthClientFunction(this.chatConfig)
            });
            
            TelemetryHelper.logFacadeChatSDKEventToAllTelemetry(LogLevel.ERROR, {
                Event: TelemetryEvent.NewTokenValidationFailed,
                Description: "GetAuthToken function is not present",
                ExceptionDetails: "Missing function : " + getAuthClientFunction(this.chatConfig)
            });

            return { result: false, message: "GetAuthToken function is not present" };
        }

        console.info("[LCW][FacadeChatSDK][tokenRing] invoking handleAuthentication", {
            using: this.getAuthToken !== undefined ? "props.getAuthToken" : "chatSDK.chatSDKConfig.getAuthToken",
            authClientFunction: getAuthClientFunction(this.chatConfig)
        });

        // if token is not set, or token is already expired , then go to grab a token
        this.token = "";
        this.expiration = 0;

        try {
            console.info("[LCW][FacadeChatSDK][tokenRing] calling handleAuthentication...");
            const ring = await handleAuthentication(this.chatSDK, this.chatConfig, this.getAuthToken);

            console.info("[LCW][FacadeChatSDK][tokenRing] handleAuthentication returned", {
                result: ring?.result,
                hasToken: !!ring?.token,
                tokenLength: ring?.token?.length || 0,
                error: ring?.error
            });

            if (ring?.result === true && ring?.token) {
                await this.setToken(ring.token);

                console.info("[LCW][FacadeChatSDK][tokenRing] SUCCESS: New Token obtained", {
                    expiration: this.expiration
                });

                TelemetryHelper.logFacadeChatSDKEventToAllTelemetry(LogLevel.INFO, {
                    Event: TelemetryEvent.NewTokenValidationCompleted,
                    Description: "New Token obtained",
                    Data: {
                        "Token_Expiration": this.expiration
                    }
                });
                return { result: true, message: "New Token obtained" };
            }

            // Mid-auth: no token available - set pending flag for startChat to handle
            const isEmptyTokenWithoutError = isNullOrEmptyString(ring?.token) && 
                (ring?.result === true || (ring?.result === false && !ring?.error));
            
            console.info("[LCW][FacadeChatSDK][tokenRing] Checking mid-auth case", {
                isMidAuthEnabled: this.isMidAuthEnabled(),
                isEmptyTokenWithoutError,
                ringResult: ring?.result,
                hasRingError: !!ring?.error
            });

            if (this.isMidAuthEnabled() && isEmptyTokenWithoutError) {
                console.info("[LCW][FacadeChatSDK][tokenRing] Mid-auth enabled and no token returned - proceeding as unauthenticated", {
                    result: ring?.result,
                    hasError: !!ring?.error
                });

                // Clear both Facade and SDK token state so API calls use UNAUTHENTICATED endpoints
                // This is important for getConversationDetails to work on unauthenticated chats
                this.token = "";
                this.expiration = 0;
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                (this.chatSDK as any).authenticatedUserToken = null;
                
                // Set flag for startChat to handle the state transition appropriately
                this.pendingMidAuthUnauthenticatedState = true;

                console.info("[LCW][FacadeChatSDK][tokenRing] Cleared auth state for unauthenticated flow", {
                    facadeTokenSet: !!this.token,
                });

                // Log as INFO since this is expected behavior for mid-auth
                TelemetryHelper.logFacadeChatSDKEventToAllTelemetry(LogLevel.INFO, {
                    Event: TelemetryEvent.NewTokenValidationCompleted,
                    Description: "Mid-auth enabled: no token returned; proceeding as unauthenticated"
                });

                return { result: true, message: "Mid-auth: proceeding as unauthenticated" };
            }

            console.info("[LCW][FacadeChatSDK][tokenRing] FAILED: handleAuthentication did not return token", {
                result: ring?.result,
                errorMessage: ring?.error?.message
            });
            TelemetryHelper.logFacadeChatSDKEventToAllTelemetry(LogLevel.ERROR, {
                Event: TelemetryEvent.NewTokenValidationFailed,
                Description: ring.error?.message,
                ExceptionDetails: ring?.error
            });
            return {
                result: false,
                message: ring?.error?.message || "Failed to get token"
            };
        } catch (e: unknown) {
            console.error("Unexpected error while getting token", e);
            TelemetryHelper.logFacadeChatSDKEventToAllTelemetry(LogLevel.ERROR, {
                Event: TelemetryEvent.NewTokenValidationFailed,
                Description: "Unexpected error while getting token",
                ExceptionDetails: e
            });
            return { result: false, message: "Unexpected error while getting token" };
        }
    }

    /**
     * Sets the state for mid-auth unauthenticated flow.
     * Called when mid-auth is enabled but no token is available.
     */
    private setMidAuthUnauthenticatedState(): void {
        console.info("[LCW][FacadeChatSDK][setMidAuthUnauthenticatedState] Setting up unauthenticated state for mid-auth");
        
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const sdk = this.chatSDK as any;
        const hadExistingChat = !!sdk.chatToken?.chatId;
        const previousChatId = sdk.chatToken?.chatId;
        
        console.info("[LCW][FacadeChatSDK][setMidAuthUnauthenticatedState] Current SDK state before clearing", {
            hasChatToken: hadExistingChat,
            chatId: previousChatId,
            hasReconnectId: !!sdk.reconnectId,
            hasRequestId: !!sdk.requestId
        });
        
        // Clear FacadeChatSDK auth state
        this.clearAuthState();
        
        // Clear SDK internal state to prevent reconnection to previous session
        sdk.chatToken = {};
        sdk.reconnectId = null;
        sdk.requestId = null;
        sdk.sessionId = null;
        sdk.conversation = null;
        
        console.info("[LCW][FacadeChatSDK][setMidAuthUnauthenticatedState] SDK state cleared for fresh unauthenticated chat");
        
        if (hadExistingChat) {
            TelemetryHelper.logFacadeChatSDKEventToAllTelemetry(LogLevel.INFO, {
                Event: TelemetryEvent.MidConversationAuthReset,
                Description: "Mid-auth without token: local state cleared",
                Data: { previousChatId }
            });
        }
    }

    /**
     * Clears authentication state in both FacadeChatSDK and underlying SDK
     */
    private clearAuthState(): void {
        console.info("[LCW][FacadeChatSDK][clearAuthState] Clearing auth state");
        this.token = "";
        this.expiration = 0;
        this.isAuthenticated = false;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (this.chatSDK as any).authenticatedUserToken = null;
    }

    /**
     * MID-AUTH: Migrate conversation from unauthenticated to authenticated.
     * Called after startChat() when user has a valid token but the backend conversation
     * was started as unauthenticated (needs migration).
     * 
     * @param isReconnect - Whether this is a reconnect scenario
     * @returns Promise that resolves when migration is complete (or fails gracefully)
     */
    private async migrateConversationToAuthenticated(isReconnect: boolean): Promise<void> {
        console.info("[LCW][FacadeChatSDK][migrateConversationToAuthenticated] Calling authenticateChat to migrate conversation", {
            isReconnect
        });
        
        try {
            await this.chatSDK.authenticateChat(this.token as string, { refreshChatToken: true });
            
            // Update Facade's auth state to reflect successful authentication
            this.isAuthenticated = true;
            
            console.info("[LCW][FacadeChatSDK][migrateConversationToAuthenticated] authenticateChat completed - conversation now authenticated");
            
            TelemetryHelper.logFacadeChatSDKEventToAllTelemetry(LogLevel.INFO, {
                Event: TelemetryEvent.MidConversationAuthSucceeded,
                Description: "Mid-auth: authenticateChat completed, conversation migrated to authenticated"
            });
        } catch (e) {
            // Non-fatal: Chat is already active via startChat
            // Log warning but don't fail the startChat operation
            console.warn("[LCW][FacadeChatSDK][migrateConversationToAuthenticated] authenticateChat returned error (chat still active, will retry on next reconnect)", e);
            TelemetryHelper.logFacadeChatSDKEventToAllTelemetry(LogLevel.WARN, {
                Event: TelemetryEvent.MidConversationAuthFailed,
                Description: "Mid-auth: authenticateChat returned error after startChat, chat still active",
                ExceptionDetails: { message: (e as Error)?.message }
            });
        }
    }

    /**
     * MID-AUTH: Configure SDK auth state before calling startChat.
     * Handles two cases:
     * - CASE 1: Pending unauthenticated state (tokenRing couldn't get token - user not logged in)
     * - CASE 2: Authenticated with valid token (tokenRing got token - user is logged in)
     *
     * Note: The "user logged in mid-conversation" scenario is now handled directly in tokenRing(),
     * which runs on every SDK call and can detect login without requiring a new startChat call.
     *
     * @returns Object containing updated state for startChat
     */
    private async configureMidAuthState(
        isReconnect: boolean,
        wasPreviousSessionAuthenticated: boolean
    ): Promise<{ shouldClearReconnectParams: boolean }> {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const sdk = this.chatSDK as any;

        console.info("[LCW][FacadeChatSDK][configureMidAuthState] Checking auth state", {
            pendingMidAuthUnauthenticatedState: this.pendingMidAuthUnauthenticatedState,
            isAuthenticated: this.isAuthenticated,
            hasToken: this.isTokenSet(),
            isReconnect,
            wasPreviousSessionAuthenticated
        });

        // CASE 1: Pending unauthenticated state (from tokenRing - no token available)
        // Note: pendingMidAuthUnauthenticatedState is NOT reset here. It stays true until the user
        // actually logs in (cleared in tokenRing when token is obtained). This ensures:
        // 1. tokenRing checks for login on every SDK call (mid-conversation detection)
        // 2. CASE 1 re-triggers on every startChat to set deferInitialAuth = true
        if (this.pendingMidAuthUnauthenticatedState) {
            const shouldClear = this.handlePendingUnauthenticatedState(wasPreviousSessionAuthenticated);
            sdk.deferInitialAuth = true;
            return { shouldClearReconnectParams: shouldClear };
        }

        // CASE 2: Authenticated with valid token (tokenRing got token or upgraded mid-conversation)
        if (this.isTokenSet() && !this.isTokenExpired()) {
            this.handleAuthenticatedState(isReconnect, wasPreviousSessionAuthenticated);
        }

        return { shouldClearReconnectParams: false };
    }

    /**
     * Handle CASE 1: Pending unauthenticated state from tokenRing
     * @returns true if reconnect params should be cleared (Auth → Unauth transition)
     */
    private handlePendingUnauthenticatedState(wasPreviousSessionAuthenticated: boolean): boolean {
        if (wasPreviousSessionAuthenticated) {
            // Auth → Unauth transition: user logged out, start fresh
            console.info("[LCW][FacadeChatSDK][handlePendingUnauthenticatedState] Auth → Unauth: clearing state for new chat");
            this.setMidAuthUnauthenticatedState();
            return true;
        }
        
        // Unauth → Unauth: keep liveChatContext for reconnection
        console.info("[LCW][FacadeChatSDK][handlePendingUnauthenticatedState] Unauth → Unauth: keeping liveChatContext");
        this.isAuthenticated = false;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (this.chatSDK as any).authenticatedUserToken = null;
        return false;
    }

    /**
     * Handle CASE 2: Authenticated with valid token
     * Only set deferInitialAuth for reconnects to unauthenticated sessions (need migration)
     * For new chats or reconnects to authenticated sessions, SDK handles auth internally
     */
    private handleAuthenticatedState(
        isReconnect: boolean,
        wasPreviousSessionAuthenticated: boolean
    ): void {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const sdk = this.chatSDK as any;
        
        console.info("[LCW][FacadeChatSDK][handleAuthenticatedState] Authenticated flow with valid token", {
            isReconnect,
            wasPreviousSessionAuthenticated
        });
        
        sdk.authenticatedUserToken = this.token;

        // Defer auth only for reconnects to unauthenticated sessions (will call authenticateChat after)
        if (isReconnect && !wasPreviousSessionAuthenticated) {
            sdk.deferInitialAuth = true;
        } else {
            // Explicitly reset for new chats or reconnects to authenticated sessions
            // Prevents inheriting deferInitialAuth=true from a previous unauthenticated chat
            sdk.deferInitialAuth = false;
        }
    }

    private async validateAndExecuteCall<T>(functionName: string, fn: () => Promise<T>): Promise<T> {
        console.info(`[LCW][FacadeChatSDK][${functionName}] validateAndExecuteCall START`);
        const pingResponse = await this.tokenRing();
        
        console.info(`[LCW][FacadeChatSDK][${functionName}] tokenRing returned`, {
            result: pingResponse.result,
            message: pingResponse.message
        });

        if (pingResponse.result === true) {
            console.info(`[LCW][FacadeChatSDK][${functionName}] Executing SDK function...`);
            return fn();
        }

        const executionErrorMessage = "Authentication Setup Error: Token validation failed - GetAuthToken function is not present";
        //telemetry is already logged in tokenRing, so no need to log again, just return the error and communicate to the console
        console.error(`${executionErrorMessage} Additional details: Process to get a token failed for ${functionName}, ${pingResponse.message}`);
        BroadcastService.postMessage({
            eventName: BroadcastEvent.OnWidgetError,
            payload: {
                errorMessage: executionErrorMessage,
            }
        });
        throw new Error(executionErrorMessage);
    }

    public async initialize(optionalParams: InitializeOptionalParams = {}): Promise<ChatConfig> {
        return this.validateAndExecuteCall("initialize", () => this.chatSDK.initialize(optionalParams));
    }

    public async getChatReconnectContext(optionalParams: ChatReconnectOptionalParams = {}): Promise<ChatReconnectContext> {
        return this.validateAndExecuteCall("getChatReconnectContext", () => this.chatSDK.getChatReconnectContext(optionalParams));
    }

    public async startChat(optionalParams: StartChatOptionalParams = {}): Promise<void> {
        const midAuthEnabled = this.isMidAuthEnabled();
        const isReconnect = !!optionalParams.liveChatContext || !!optionalParams.reconnectId;
        const wasPreviousSessionAuthenticated = optionalParams.wasAuthenticated === true;
        
        console.info("[LCW][FacadeChatSDK][startChat] START", {
            isAuthenticated: this.isAuthenticated,
            hasToken: this.isTokenSet(),
            isMidAuthEnabled: midAuthEnabled,
            isReconnect,
            wasPreviousSessionAuthenticated
        });
        
        return this.validateAndExecuteCall("startChat", async () => {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const sdk = this.chatSDK as any;
            
            // Configure mid-auth state if enabled (otherwise existing behavior)
            if (midAuthEnabled) {

                const { shouldClearReconnectParams } = await this.configureMidAuthState(
                    isReconnect,
                    wasPreviousSessionAuthenticated
                );
                
                // Clear reconnect params if transitioning from Auth → Unauth
                if (shouldClearReconnectParams) {
                    delete optionalParams.liveChatContext;
                    delete optionalParams.reconnectId;
                }
            }
            
            console.info("[LCW][FacadeChatSDK][startChat] Calling SDK startChat", {
                isAuthenticated: this.isAuthenticated,
                deferInitialAuth: sdk.deferInitialAuth,
                hasLiveChatContext: !!optionalParams.liveChatContext,
                hasReconnectId: !!optionalParams.reconnectId
            });
            
            await this.chatSDK.startChat(optionalParams);
            
            // Migrate to authenticated if needed (only for reconnects to unauthenticated sessions)
            // New chats don't need migration - SDK handles auth internally during startChat
            const shouldMigrateToAuth = midAuthEnabled && 
                                        isReconnect &&
                                        this.isTokenSet() && 
                                        !this.isTokenExpired() && 
                                        !wasPreviousSessionAuthenticated;
            
            if (shouldMigrateToAuth) {
                await this.migrateConversationToAuthenticated(isReconnect);
            }
            
            // SINGLE SOURCE OF TRUTH: Broadcast final auth state after startChat completes
            // Only broadcast when auth state CHANGES to avoid unnecessary updates
            if (midAuthEnabled) {
                const isAuthenticatedAfterStart = this.isTokenSet() && !this.isTokenExpired();
                
                // Only broadcast if state changed:
                // - New chats: always broadcast (widget needs initial state)
                // - Reconnects: only if auth state differs from previous session
                const authStateChanged = !isReconnect || (isAuthenticatedAfterStart !== wasPreviousSessionAuthenticated);
                
                if (authStateChanged) {
                    console.info("[LCW][FacadeChatSDK][startChat] Auth state changed, broadcasting", {
                        isAuthenticated: isAuthenticatedAfterStart,
                        isReconnect,
                        wasPreviousSessionAuthenticated
                    });
                    
                    BroadcastService.postMessage({
                        eventName: isAuthenticatedAfterStart 
                            ? BroadcastEvent.MidConversationAuthSucceeded 
                            : BroadcastEvent.MidConversationAuthReset,
                        payload: { 
                            isAuthenticated: isAuthenticatedAfterStart,
                            isStartChatComplete: true,
                            isReconnect
                        }
                    });
                } else {
                    console.info("[LCW][FacadeChatSDK][startChat] Auth state unchanged, skipping broadcast", {
                        isAuthenticated: isAuthenticatedAfterStart,
                        isReconnect
                    });
                }
            }
        });
    }

    public async endChat(optionalParams: EndChatOptionalParams = {}): Promise<void> {
        return this.validateAndExecuteCall("endChat", () => this.chatSDK.endChat(optionalParams));
    }

    public async getCurrentLiveChatContext(): Promise<object> {
        return this.validateAndExecuteCall("getCurrentLiveChatContext", () => this.chatSDK.getCurrentLiveChatContext());
    }

    public async getConversationDetails(optionalParams: GetConversationDetailsOptionalParams = {}): Promise<LiveWorkItemDetails> {
        return this.validateAndExecuteCall("getConversationDetails", () => this.chatSDK.getConversationDetails(optionalParams));
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public async getPreChatSurvey(parse = true): Promise<any> {
        //prechat survey is obtained from config object, which is not required to be authenticated
        // removing the tokenRing function from this call for backward compatibility
        // TODO ::  wrap this function around authentication
        return this.chatSDK.getPreChatSurvey(parse);
    }

    public async getLiveChatConfig(optionalParams?: GetLiveChatConfigOptionalParams): Promise<ChatConfig> {
        return this.validateAndExecuteCall("getLiveChatConfig", () => this.chatSDK.getLiveChatConfig(optionalParams));
    }

    public async getChatToken(cached = true, optionalParams?: GetChatTokenOptionalParams): Promise<IChatToken> {
        return this.validateAndExecuteCall("getChatToken", () => this.chatSDK.getChatToken(cached, optionalParams));
    }

    public async getCallingToken(): Promise<string> {
        return this.validateAndExecuteCall("getCallingToken", () => this.chatSDK.getCallingToken());
    }

    public async getMessages(): Promise<IMessage[] | OmnichannelMessage[] | undefined> {
        return this.validateAndExecuteCall("getMessages", () => this.chatSDK.getMessages());
    }

    public async getDataMaskingRules(): Promise<MaskingRules> {
        return this.validateAndExecuteCall("getDataMaskingRules", () => this.chatSDK.getDataMaskingRules());
    }

    public async sendMessage(message: ChatSDKMessage): Promise<void | OmnichannelMessage> {
        return this.validateAndExecuteCall("sendMessage", () => this.chatSDK.sendMessage(message));
    }

    public async onNewMessage(onNewMessageCallback: CallableFunction, optionalParams: OnNewMessageOptionalParams = { disablePolling: false }): Promise<void> {
        return this.validateAndExecuteCall("onNewMessage", () => this.chatSDK.onNewMessage(onNewMessageCallback, optionalParams));
    }

    public async sendTypingEvent(): Promise<void> {
        return this.validateAndExecuteCall("sendTypingEvent", () => this.chatSDK.sendTypingEvent());
    }

    public async onTypingEvent(onTypingEventCallback: CallableFunction): Promise<void> {
        return this.validateAndExecuteCall("onTypingEvent", () => this.chatSDK.onTypingEvent(onTypingEventCallback));
    }

    public async onAgentEndSession(onAgentEndSessionCallback: (message: IRawThread | ParticipantsRemovedEvent) => void): Promise<void> {
        return this.validateAndExecuteCall("onAgentEndSession", () => this.chatSDK.onAgentEndSession(onAgentEndSessionCallback));
    }

    public async uploadFileAttachment(fileInfo: IFileInfo | File): Promise<IRawMessage | OmnichannelMessage> {
        return this.validateAndExecuteCall("uploadFileAttachment", () => this.chatSDK.uploadFileAttachment(fileInfo));
    }

    public async downloadFileAttachment(fileMetadata: FileMetadata | IFileMetadata): Promise<Blob> {
        return this.validateAndExecuteCall("downloadFileAttachment", () => this.chatSDK.downloadFileAttachment(fileMetadata));
    }

    public async emailLiveChatTranscript(body: ChatTranscriptBody, optionalParams: EmailLiveChatTranscriptOptionaParams = {}): Promise<void> {
        return this.validateAndExecuteCall("emailLiveChatTranscript", () => this.chatSDK.emailLiveChatTranscript(body, optionalParams));
    }

    public async getLiveChatTranscript(optionalParams: GetLiveChatTranscriptOptionalParams = {}): Promise<GetLiveChatTranscriptResponse> {
        return this.validateAndExecuteCall("getLiveChatTranscript", () => this.chatSDK.getLiveChatTranscript(optionalParams));
    }

    // response from origin is unknown, but this definition breaks create adapter for shimAdapter, switching to any until type is returned from origin
    public async createChatAdapter(optionalParams: ChatAdapterOptionalParams = {}): Promise<ChatAdapter> {
        return this.validateAndExecuteCall("createChatAdapter", () => this.chatSDK.createChatAdapter(optionalParams));
    }

    public async isVoiceVideoCallingEnabled(): Promise<boolean> {
        this.tokenRing();
        return this.chatSDK.isVoiceVideoCallingEnabled();
    }

    public async getVoiceVideoCalling(params: VoiceVideoCallingOptionalParams = {}): Promise<GetVoiceVideoCallingResponse> {
        return this.validateAndExecuteCall("getVoiceVideoCalling", () => this.chatSDK.getVoiceVideoCalling(params));
    }

    public async getPostChatSurveyContext(): Promise<PostChatContext> {
        return this.validateAndExecuteCall("getPostChatSurveyContext", () => this.chatSDK.getPostChatSurveyContext());
    }

    public async getAgentAvailability(optionalParams: GetAgentAvailabilityOptionalParams = {}): Promise<GetAgentAvailabilityResponse> {
        return this.validateAndExecuteCall("getAgentAvailability", () => this.chatSDK.getAgentAvailability(optionalParams));
    }
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public async getReconnectableChats(reconnectableChatsParams: any = {}): Promise<any> {

        /**
         * 
         * This is a particular case, we dont expose getReconnectableChats in the SDK,
         * The only way to use is by tunneling directly from the SDK to OCSDK,
         * 
         * In case of prechat, the function is called before any formal authentication is made, 
         * this is an specific case for persistent chats, to prevent the survey be loaded again for an on going chat,
         * 
         * In this case, we check for existance of the token , otherwise we perform the authentication, error is propagated in case of issues.
         * 
         * Once the token is obtained , this will be added to the params to call the function.
         * 
         * This is a particular case, should not be taken as pattern.
                    }
         *
         */

        if (this.token === null || this.token === "") {
            // If token is not set, try to get it using tokenRing
            const pingResponse = await this.tokenRing();
            if (pingResponse.result === false) {
                const errorMessage = "Authentication Setup Error: Token validation failed for reconnectable chats";
                //telemetry is already logged in tokenRing, so no need to log again, just return the error and communicate to the console
                console.error(`Authentication failed: Process to get a token failed for getReconnectableChats, ${pingResponse.message}`);
                BroadcastService.postMessage({
                    eventName: BroadcastEvent.OnWidgetError,
                    payload: {
                        errorMessage: errorMessage,
                    }
                });
                throw new Error(errorMessage);
            }
        }
        
        // Always override the token in params regardless of how getReconnectableChats was called
        reconnectableChatsParams.authenticatedUserToken = this.token;
        
        return this.validateAndExecuteCall("getReconnectableChats", () => this.chatSDK.OCClient.getReconnectableChats(reconnectableChatsParams));

    }

    public async fetchPersistentConversationHistory(getPersistentChatHistoryOptionalParams: GetPersistentChatHistoryOptionalParams = {}): Promise<GetPersistentChatHistoryResponse> {
        return this.validateAndExecuteCall("getPersistentChatHistory", () => this.chatSDK.getPersistentChatHistory(getPersistentChatHistoryOptionalParams));
    }
}