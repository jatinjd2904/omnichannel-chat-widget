import { FacadeChatSDK } from "./FacadeChatSDK";
import { IFacadeChatSDKInput } from "./types/IFacadeChatSDKInput";
import { OmnichannelChatSDK } from "@microsoft/omnichannel-chat-sdk";
import { handleAuthentication } from "../../components/livechatwidget/common/authHelper";
import { BroadcastService } from "@microsoft/omnichannel-chat-components";
import { BroadcastEvent } from "../telemetry/TelemetryConstants";

// mock BroadcastService
jest.mock("@microsoft/omnichannel-chat-components", () =>({
    BroadcastService: {
        postMessage: jest.fn()
    }}));

jest.mock("../../components/livechatwidget/common/authHelper");
// function to mimic a jwt token with exp time from now to 5 min in the future in seconds
function getJWTToken() {
    const now = Math.floor(Date.now() / 1000);
    const expiration = now + 300; // 5 minutes in the future (300 seconds)
    const tokenPayload = {
        sub: "user_id+2",
        name: "hey@hey.com",
        data: "http://example.com/?id=123-45\\456&name=John+Doe&age=30",
        special: "+/=?\\",
        html: "<div><h1>Hello, World!</h1><p>This is a <strong>test</strong> HTML snippet.</p></div>",
        lwicontexts: "{\"BingConsent\":{\"value\":\"true\",\"isDisplayable\":true},\"HostingApp\":{\"value\":\"M365AdminPortal\",\"isDisplayable\":true},\"ClientContext0\":{\"value\":\"{\\\"profile\\\":{\\\"billing\\\":{\\\"list\\\":{\\\"Subscriptions\\\":[{\\\"NextSubscriptionState\\\":\\\"InGracePeriod\\\",\\\"Channel\\\":\\\"Direct\\\",\\\"IsLifecycleExpired\\\":false,\\\"TotalLicenseQuantity\\\":1,\\\"IsInGracePeriod\\\":false,\\\"Name\\\":null,\\\"SubscriptionEndDate\\\":\\\"2025-08-08T00:00:00Z\\\",\\\"SubscriptionStartDate\\\":\\\"2023-08-08T00:00:00Z\\\",\\\"SubscriptionState\\\":1,\"isDisplayable\":true}}",
        exp: expiration
    };
    const tokenPayloadBase64 = Buffer.from(JSON.stringify(tokenPayload)).toString("base64");
    return {
        token: `eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.${tokenPayloadBase64}.signature`,
        expiration: expiration
    };
}

function getJWTTokenBase64Url() {
    const now = Math.floor(Date.now() / 1000);
    const expiration = now + 300; // 5 minutes in the future (300 seconds)
    const tokenPayload = {
        sub: "user_id+2",
        name: "hey@hey.com",
        data: "http://example.com/?id=123-45\\456&name=John+Doe&age=30",
        special: "+/=?\\",
        html: "<div><h1>Hello, World!</h1><p>This is a <strong>test</strong> HTML snippet.</p></div>",
        lwicontexts: "{\"BingConsent\":{\"value\":\"true\",\"isDisplayable\":true},\"HostingApp\":{\"value\":\"M365AdminPortal\",\"isDisplayable\":true},\"ClientContext0\":{\"value\":\"{\\\"profile\\\":{\\\"billing\\\":{\\\"list\\\":{\\\"Subscriptions\\\":[{\\\"NextSubscriptionState\\\":\\\"InGracePeriod\\\",\\\"Channel\\\":\\\"Direct\\\",\\\"IsLifecycleExpired\\\":false,\\\"TotalLicenseQuantity\\\":1,\\\"IsInGracePeriod\\\":false,\\\"Name\\\":null,\\\"SubscriptionEndDate\\\":\\\"2025-08-08T00:00:00Z\\\",\\\"SubscriptionStartDate\\\":\\\"2023-08-08T00:00:00Z\\\",\\\"SubscriptionState\\\":1,\"isDisplayable\":true}}",
        exp: expiration
    };

    // Convert payload to JSON string and then to Base64URL
    const tokenPayloadBase64 = Buffer.from(JSON.stringify(tokenPayload)).toString("base64");
    const base64url = tokenPayloadBase64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
    return {
        token: `eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.${base64url}.signature`,
        expiration: expiration
    };
}

describe("FacadeChatSDK", () => {
    let facadeChatSDK: FacadeChatSDK;

    beforeEach(() => {
        const input: IFacadeChatSDKInput = {
            chatSDK: new OmnichannelChatSDK({
                orgId: "your-org-id",
                orgUrl: "https://your-org-url",
                widgetId: "your-widget-id"
            }),
            chatConfig: {
                ChatWidgetLanguage: undefined,
                DataMaskingInfo: undefined,
                LiveChatConfigAuthSettings: undefined,
                LiveChatVersion: 0,
                LiveWSAndLiveChatEngJoin: undefined,
                allowedFileExtensions: "",
                maxUploadFileSize: ""
            },
            getAuthToken: jest.fn(),
            isAuthenticated: true,
            isSDKMocked: false
        };
        facadeChatSDK = new FacadeChatSDK(input, false);
    });

    describe("convertExpiration", () => {
        it("should convert milliseconds to seconds", () => {
            const expiration = 1609459200000; // 2021-01-01T00:00:00.000Z in milliseconds
            const result = facadeChatSDK["convertExpiration"](expiration);
            expect(result).toBe(1609459200); // 2021-01-01T00:00:00Z in seconds
        });

        it("should return seconds as is", () => {
            const expiration = 1609459200; // 2021-01-01T00:00:00Z in seconds
            const result = facadeChatSDK["convertExpiration"](expiration);
            expect(result).toBe(1609459200);
        });

        it("should return 0 if expiration is undefined", () => {
            const result = facadeChatSDK["convertExpiration"](undefined);
            expect(result).toBe(0);
        });
    });

    describe("isTokenExpired", () => {
        it("should return false if expiration is 0", () => {
            facadeChatSDK["expiration"] = 0;
            const result = facadeChatSDK["isTokenExpired"]();
            expect(result).toBe(false);
        });

        it("should return true if current time is greater than expiration", () => {
            const now = Math.floor(Date.now() / 1000);
            facadeChatSDK["expiration"] = now - 10; // expired 10 seconds ago
            const result = facadeChatSDK["isTokenExpired"]();
            expect(result).toBe(true);
        });

        it("should return false if current time is less than expiration", () => {
            const now = Math.floor(Date.now() / 1000);
            facadeChatSDK["expiration"] = now + 10; // expires in 10 seconds
            const result = facadeChatSDK["isTokenExpired"]();
            expect(result).toBe(false);
        });
    });

    describe("setToken", () => {
        it("should set token and expiration correctly", async () => {
            const jwt = getJWTToken();
            await facadeChatSDK["setToken"](jwt.token);
            expect(facadeChatSDK["token"]).toBe(jwt.token);
            expect(facadeChatSDK["expiration"]).toBe(jwt.expiration);
        });

        it("Base64url JWT should be valid", async () => {
            const jwt = getJWTTokenBase64Url();
            await facadeChatSDK["setToken"](jwt.token);
            expect(facadeChatSDK["token"]).toBe(jwt.token);
            expect(facadeChatSDK["expiration"]).toBe(jwt.expiration);
           
        });

        it("should throw error for invalid token format", async () => {
            const token = "any-thing";
            await expect(facadeChatSDK["setToken"](token)).rejects.toThrow("Authentication Setup Error: Invalid token format, must be in JWT format");
        });

        it("should throw error for expired token", async () => {
            const token = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJleHAiOjE2MDk0NTkyMDB9.4f1d4f1d4f1d4f1d4f1d4f1d4f1d4f1d4f1d4f1d4f1d4f1d4f1d4f1d4f1d";
            jest.spyOn(Date, "now").mockReturnValue(1609459201000); // mock current time to be after token expiration
            await expect(facadeChatSDK["setToken"](token)).rejects.toThrow("Authentication Setup Error: New authentication token is already expired");
        });

    });

    describe("tokenRing", () => {
        it("should return true if sdkMocked is true", async () => {
            facadeChatSDK["sdkMocked"] = true;
            const result = await facadeChatSDK["tokenRing"]();
            expect(result).toEqual({ result: true, message: "Authentication not needed" });
        });

        it("should return true if isAuthenticated is false", async () => {
            facadeChatSDK["isAuthenticated"] = false;
            const result = await facadeChatSDK["tokenRing"]();
            expect(result).toEqual({ result: true, message: "Authentication not needed" });
        });

        it("should return true if token is valid", async () => {
            facadeChatSDK["token"] = "valid.token";
            facadeChatSDK["expiration"] = Math.floor(Date.now() / 1000) + 1000; // valid token
            const result = await facadeChatSDK["tokenRing"]();
            expect(result).toEqual({ result: true, message: "Token is valid" });
        });

        it("should return true if Facade is disabled", async () => {

            const input: IFacadeChatSDKInput = {
                chatSDK: new OmnichannelChatSDK({
                    orgId: "your-org-id",
                    orgUrl: "https://your-org-url",
                    widgetId: "your-widget-id"
                }),
                chatConfig: {
                    ChatWidgetLanguage: undefined,
                    DataMaskingInfo: undefined,
                    LiveChatConfigAuthSettings: undefined,
                    LiveChatVersion: 0,
                    LiveWSAndLiveChatEngJoin: undefined,
                    allowedFileExtensions: "",
                    maxUploadFileSize: ""
                },
                getAuthToken: jest.fn(),
                isAuthenticated: true,
                isSDKMocked: false
            };
            facadeChatSDK = new FacadeChatSDK(input, true);

            const result = await facadeChatSDK["tokenRing"]();
            expect(result).toEqual({ result: true, message: "Facade is disabled" });
        });


        it("should return false if getAuthToken is undefined", async () => {
            facadeChatSDK["getAuthToken"] = undefined;
            const result = await facadeChatSDK["tokenRing"]();
            expect(result).toEqual({ result: false, message: "GetAuthToken function is not present" });
        });

        it("should set new token if obtained successfully", async () => {
            const newToken = getJWTToken();
            (handleAuthentication as jest.Mock).mockResolvedValue({ result: true, token: newToken.token });
            const result = await facadeChatSDK["tokenRing"]();
            expect(result).toEqual({ result: true, message: "New Token obtained" });
            expect(facadeChatSDK["token"]).toBe(newToken.token);
        });

    });

    describe("validateAndExecuteCall", () => {
        it("should execute function if token is valid", async () => {
            jest.spyOn(facadeChatSDK, "tokenRing").mockResolvedValue({ result: true, message: "Token is valid" });
            const mockFn = jest.fn().mockResolvedValue("success");
            const result = await facadeChatSDK["validateAndExecuteCall"]("testFunction", mockFn);
            expect(result).toBe("success");
        });

        it("should throw error if token is invalid", async () => {
            jest.spyOn(facadeChatSDK, "tokenRing").mockResolvedValue({ result: false, message: "Token is invalid" });
            const mockFn = jest.fn();
            await expect(facadeChatSDK["validateAndExecuteCall"]("testFunction", mockFn)).rejects.toThrow("Authentication Setup Error: Token validation failed - GetAuthToken function is not present");
        });
    });

    describe("isMidAuthEnabled", () => {
        it("should return true when msdyn_authenticatedsigninoptional is 'true'", () => {
            facadeChatSDK["chatConfig"] = {
                LiveWSAndLiveChatEngJoin: {
                    msdyn_authenticatedsigninoptional: "true"
                }
            } as unknown;
            expect(facadeChatSDK["isMidAuthEnabled"]()).toBe(true);
        });

        it("should return true when msdyn_authenticatedsigninoptional is 'True' (case insensitive)", () => {
            facadeChatSDK["chatConfig"] = {
                LiveWSAndLiveChatEngJoin: {
                    msdyn_authenticatedsigninoptional: "True"
                }
            } as unknown;
            expect(facadeChatSDK["isMidAuthEnabled"]()).toBe(true);
        });

        it("should return false when msdyn_authenticatedsigninoptional is 'false'", () => {
            facadeChatSDK["chatConfig"] = {
                LiveWSAndLiveChatEngJoin: {
                    msdyn_authenticatedsigninoptional: "false"
                }
            } as unknown;
            expect(facadeChatSDK["isMidAuthEnabled"]()).toBe(false);
        });

        it("should return false when msdyn_authenticatedsigninoptional is undefined", () => {
            facadeChatSDK["chatConfig"] = {
                LiveWSAndLiveChatEngJoin: {}
            } as unknown;
            expect(facadeChatSDK["isMidAuthEnabled"]()).toBe(false);
        });

        it("should return false when LiveWSAndLiveChatEngJoin is undefined", () => {
            facadeChatSDK["chatConfig"] = {} as unknown;
            expect(facadeChatSDK["isMidAuthEnabled"]()).toBe(false);
        });
    });

    describe("isTokenExpiredByValue", () => {
        it("should return true for empty string token", () => {
            const result = facadeChatSDK["isTokenExpiredByValue"]("");
            expect(result).toBe(true);
        });

        it("should return true for null token", () => {
            const result = facadeChatSDK["isTokenExpiredByValue"](null as unknown);
            expect(result).toBe(true);
        });

        it("should return false for valid non-expired token", () => {
            const jwt = getJWTToken();
            const result = facadeChatSDK["isTokenExpiredByValue"](jwt.token);
            expect(result).toBe(false);
        });

        it("should return true for expired token", () => {
            // Create an expired token
            const expiredPayload = { exp: Math.floor(Date.now() / 1000) - 300 }; // 5 minutes ago
            const payloadBase64 = Buffer.from(JSON.stringify(expiredPayload)).toString("base64");
            const expiredToken = `eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.${payloadBase64}.signature`;
            const result = facadeChatSDK["isTokenExpiredByValue"](expiredToken);
            expect(result).toBe(true);
        });

        it("should return false for token without expiration (exp=0)", () => {
            const noExpPayload = { sub: "user123" }; // No exp field
            const payloadBase64 = Buffer.from(JSON.stringify(noExpPayload)).toString("base64");
            const tokenWithoutExp = `eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.${payloadBase64}.signature`;
            const result = facadeChatSDK["isTokenExpiredByValue"](tokenWithoutExp);
            expect(result).toBe(false);
        });

        it("should return true for invalid/malformed token", () => {
            const result = facadeChatSDK["isTokenExpiredByValue"]("invalid-token");
            expect(result).toBe(true);
        });
    });

    describe("setMidAuthUnauthenticatedState", () => {
        it("should clear authentication state for mid-auth unauthenticated flow", () => {
            // Set up initial authenticated state
            facadeChatSDK["token"] = "some-token";
            facadeChatSDK["expiration"] = 12345;
            facadeChatSDK["isAuthenticated"] = true;
            
            const mockChatSDK = facadeChatSDK["chatSDK"] as unknown;
            mockChatSDK.authenticatedUserToken = "some-token";
            mockChatSDK.chatToken = { chatId: "test-chat-id" };
            mockChatSDK.reconnectId = "reconnect-123";

            facadeChatSDK["setMidAuthUnauthenticatedState"]();

            expect(facadeChatSDK["token"]).toBe("");
            expect(facadeChatSDK["expiration"]).toBe(0);
            expect(facadeChatSDK["isAuthenticated"]).toBe(false);
            expect(mockChatSDK.authenticatedUserToken).toBeNull();
            expect(mockChatSDK.chatToken).toEqual({});
            expect(mockChatSDK.reconnectId).toBeNull();
        });

        it("should clear SDK internal state (chatToken, reconnectId, requestId, sessionId, conversation)", () => {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const mockChatSDK = facadeChatSDK["chatSDK"] as any;
            mockChatSDK.chatToken = { chatId: "test-chat-id" };
            mockChatSDK.reconnectId = "reconnect-123";
            mockChatSDK.requestId = "request-123";
            mockChatSDK.sessionId = "session-123";
            mockChatSDK.conversation = { disconnect: jest.fn() };

            facadeChatSDK["setMidAuthUnauthenticatedState"]();

            expect(mockChatSDK.chatToken).toEqual({});
            expect(mockChatSDK.reconnectId).toBeNull();
            expect(mockChatSDK.requestId).toBeNull();
            expect(mockChatSDK.sessionId).toBeNull();
            expect(mockChatSDK.conversation).toBeNull();
        });
    });

    describe("startChat", () => {
        let mockStartChat: jest.Mock;

        beforeEach(() => {
            // Enable mid-auth for these tests since FacadeChatSDK.startChat only
            // mutates deferInitialAuth/authenticatedUserToken when mid-auth is enabled.
            facadeChatSDK["chatConfig"] = {
                LiveWSAndLiveChatEngJoin: {
                    msdyn_authenticatedsigninoptional: "true"
                }
            } as unknown;

            mockStartChat = jest.fn().mockResolvedValue(undefined);
            facadeChatSDK["chatSDK"].startChat = mockStartChat;
            jest.spyOn(facadeChatSDK, "tokenRing").mockResolvedValue({ result: true, message: "Token is valid" });
        });

        it("should call SDK startChat with optionalParams", async () => {
            facadeChatSDK["isAuthenticated"] = false;
            
            await facadeChatSDK.startChat({ isProactiveChat: true });

            expect(mockStartChat).toHaveBeenCalled();
        });

        it("should set sdk.deferInitialAuth=true when pendingMidAuthUnauthenticatedState is true", async () => {
            facadeChatSDK["pendingMidAuthUnauthenticatedState"] = true;
            facadeChatSDK["isAuthenticated"] = false;

            await facadeChatSDK.startChat({});

            expect(mockStartChat).toHaveBeenCalled();
            // deferInitialAuth is set on the SDK instance, not on optionalParams
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const sdk = facadeChatSDK["chatSDK"] as any;
            expect(sdk.deferInitialAuth).toBe(true);
        });

        it("should clear liveChatContext when pendingMidAuthUnauthenticatedState is true and wasPreviousSessionAuthenticated", async () => {
            facadeChatSDK["pendingMidAuthUnauthenticatedState"] = true;
            facadeChatSDK["isAuthenticated"] = false;

            // Auth→Unauth transition: liveChatContext should be cleared
            const optionalParams = { liveChatContext: { chatToken: {}, requestId: "123" }, wasAuthenticated: true };
            await facadeChatSDK.startChat(optionalParams);

            expect(mockStartChat).toHaveBeenCalled();
            const callArgs = mockStartChat.mock.calls[0][0];
            expect(callArgs.liveChatContext).toBeUndefined();
        });

        it("should keep pendingMidAuthUnauthenticatedState true after startChat (not reset until token obtained)", async () => {
            facadeChatSDK["pendingMidAuthUnauthenticatedState"] = true;
            facadeChatSDK["isAuthenticated"] = false;

            await facadeChatSDK.startChat({});

            // pendingMidAuthUnauthenticatedState stays true so CASE 1 re-triggers on every startChat
            expect(facadeChatSDK["pendingMidAuthUnauthenticatedState"]).toBe(true);
        });

        it("should set authenticatedUserToken on SDK when authenticated with valid token", async () => {
            const jwt = getJWTToken();
            facadeChatSDK["isAuthenticated"] = true;
            facadeChatSDK["token"] = jwt.token;
            facadeChatSDK["expiration"] = jwt.expiration;

            await facadeChatSDK.startChat({});

            const chatSDK = facadeChatSDK["chatSDK"] as unknown;
            expect(chatSDK.authenticatedUserToken).toBe(jwt.token);
        });

        it("should set sdk.deferInitialAuth=false when authenticated with valid token for new chat", async () => {
            const jwt = getJWTToken();
            facadeChatSDK["isAuthenticated"] = true;
            facadeChatSDK["token"] = jwt.token;
            facadeChatSDK["expiration"] = jwt.expiration;

            await facadeChatSDK.startChat({});

            // deferInitialAuth is set on the SDK instance, not on optionalParams
            // For new chats with valid token, deferInitialAuth should be false
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const sdk = facadeChatSDK["chatSDK"] as any;
            expect(sdk.deferInitialAuth).toBe(false);
        });

        it("should not modify optionalParams when not authenticated", async () => {
            facadeChatSDK["isAuthenticated"] = false;
            facadeChatSDK["pendingMidAuthUnauthenticatedState"] = false;

            const optionalParams = { isProactiveChat: true };
            await facadeChatSDK.startChat(optionalParams);

            expect(mockStartChat).toHaveBeenCalledWith(optionalParams);
        });
    });

    describe("authenticateChat", () => {
        beforeEach(() => {
            jest.clearAllMocks();
        });

        describe("Pre-chat authentication (chat not started)", () => {
            beforeEach(() => {
                // Chat not started - no chatId
                (facadeChatSDK["chatSDK"] as unknown).chatToken = {};
            });

            it("should set token and isAuthenticated for pre-chat auth with string token", async () => {
                const jwt = getJWTToken();

                await facadeChatSDK.authenticateChat(jwt.token);

                expect(facadeChatSDK["token"]).toBe(jwt.token);
                expect(facadeChatSDK["isAuthenticated"]).toBe(true);
                expect((facadeChatSDK["chatSDK"] as unknown).authenticatedUserToken).toBe(jwt.token);
            });

            it("should set token and isAuthenticated for pre-chat auth with token provider function", async () => {
                const jwt = getJWTToken();
                const tokenProvider = jest.fn().mockResolvedValue(jwt.token);

                await facadeChatSDK.authenticateChat(tokenProvider);

                expect(tokenProvider).toHaveBeenCalled();
                expect(facadeChatSDK["token"]).toBe(jwt.token);
                expect(facadeChatSDK["isAuthenticated"]).toBe(true);
            });

            it("should broadcast MidConversationAuthSucceeded for pre-chat auth", async () => {
                const jwt = getJWTToken();

                await facadeChatSDK.authenticateChat(jwt.token);

                expect(BroadcastService.postMessage).toHaveBeenCalledWith({
                    eventName: BroadcastEvent.MidConversationAuthSucceeded,
                    payload: { isAuthenticated: true, token: jwt.token, isPreChatAuth: true }
                });
            });

            it("should throw error for empty token in pre-chat auth", async () => {
                await expect(facadeChatSDK.authenticateChat("")).rejects.toThrow("Authentication failed: Token is empty or null");

                expect(BroadcastService.postMessage).toHaveBeenCalledWith({
                    eventName: BroadcastEvent.OnWidgetError,
                    payload: { errorMessage: "Authentication failed: Token is empty or null" }
                });
            });

            it("should throw error for expired token in pre-chat auth", async () => {
                const expiredPayload = { exp: Math.floor(Date.now() / 1000) - 300 };
                const payloadBase64 = Buffer.from(JSON.stringify(expiredPayload)).toString("base64");
                const expiredToken = `eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.${payloadBase64}.signature`;

                await expect(facadeChatSDK.authenticateChat(expiredToken)).rejects.toThrow("Authentication Setup Error: Authentication token is already expired");
            });

            it("should clean up token and expiration on pre-chat auth failure with empty token", async () => {
                // Set some initial state
                facadeChatSDK["token"] = "old-token";
                facadeChatSDK["expiration"] = 12345;
                facadeChatSDK["isAuthenticated"] = true;

                await expect(facadeChatSDK.authenticateChat("")).rejects.toThrow();

                // Token and expiration are cleared on validation failure
                expect(facadeChatSDK["token"]).toBe("");
                expect(facadeChatSDK["expiration"]).toBe(0);
                // isAuthenticated is not changed by the empty-token validation path
                // (only the catch block for SDK failures clears it for pre-chat)
            });

            it("should throw error when token provider throws", async () => {
                const tokenProvider = jest.fn().mockRejectedValue(new Error("Provider error"));

                await expect(facadeChatSDK.authenticateChat(tokenProvider)).rejects.toThrow("Provider error");
            });
        });

        describe("Mid-conversation authentication (chat started)", () => {
            let mockAuthenticateChat: jest.Mock;

            beforeEach(() => {
                // Chat started - has chatId
                (facadeChatSDK["chatSDK"] as unknown).chatToken = { chatId: "test-chat-id" };
                mockAuthenticateChat = jest.fn().mockResolvedValue(undefined);
                facadeChatSDK["chatSDK"].authenticateChat = mockAuthenticateChat;
            });

            it("should call SDK authenticateChat for mid-conversation auth", async () => {
                const jwt = getJWTToken();

                await facadeChatSDK.authenticateChat(jwt.token);

                expect(mockAuthenticateChat).toHaveBeenCalledWith(jwt.token, {});
            });

            it("should call SDK authenticateChat with optionalParams", async () => {
                const jwt = getJWTToken();
                const optionalParams = { refreshChatToken: true };

                await facadeChatSDK.authenticateChat(jwt.token, optionalParams);

                expect(mockAuthenticateChat).toHaveBeenCalledWith(jwt.token, optionalParams);
            });

            it("should set token and isAuthenticated after mid-conversation auth", async () => {
                const jwt = getJWTToken();

                await facadeChatSDK.authenticateChat(jwt.token);

                expect(facadeChatSDK["token"]).toBe(jwt.token);
                expect(facadeChatSDK["isAuthenticated"]).toBe(true);
            });

            it("should broadcast MidConversationAuthSucceeded for mid-conversation auth", async () => {
                const jwt = getJWTToken();

                await facadeChatSDK.authenticateChat(jwt.token);

                expect(BroadcastService.postMessage).toHaveBeenCalledWith({
                    eventName: BroadcastEvent.MidConversationAuthSucceeded,
                    payload: { isAuthenticated: true, token: jwt.token, isPreChatAuth: false }
                });
            });

            it("should throw error and broadcast OnWidgetError when SDK authenticateChat fails", async () => {
                const jwt = getJWTToken();
                mockAuthenticateChat.mockRejectedValue(new Error("SDK auth failed"));

                await expect(facadeChatSDK.authenticateChat(jwt.token)).rejects.toThrow("SDK auth failed");

                expect(BroadcastService.postMessage).toHaveBeenCalledWith({
                    eventName: BroadcastEvent.OnWidgetError,
                    payload: { errorMessage: "SDK auth failed" }
                });
            });

            it("should validate token before calling SDK authenticateChat", async () => {
                await expect(facadeChatSDK.authenticateChat("")).rejects.toThrow("Authentication failed: Token is empty or null");

                expect(mockAuthenticateChat).not.toHaveBeenCalled();
            });

            it("should work with token provider function for mid-conversation auth", async () => {
                const jwt = getJWTToken();
                const tokenProvider = jest.fn().mockResolvedValue(jwt.token);

                await facadeChatSDK.authenticateChat(tokenProvider);

                expect(tokenProvider).toHaveBeenCalled();
                expect(mockAuthenticateChat).toHaveBeenCalledWith(jwt.token, {});
            });
        });
    });

    describe("tokenRing with mid-auth", () => {
        it("should set pendingMidAuthUnauthenticatedState when mid-auth enabled and no token returned", async () => {
            facadeChatSDK["chatConfig"] = {
                LiveWSAndLiveChatEngJoin: {
                    msdyn_authenticatedsigninoptional: "true"
                }
            } as unknown;
            facadeChatSDK["token"] = "";
            facadeChatSDK["expiration"] = 0;
            
            (handleAuthentication as jest.Mock).mockResolvedValue({ result: true, token: "" });

            const result = await facadeChatSDK["tokenRing"]();

            expect(result).toEqual({ result: true, message: "Mid-auth: proceeding as unauthenticated" });
            expect(facadeChatSDK["pendingMidAuthUnauthenticatedState"]).toBe(true);
        });

        it("should not set pendingMidAuthUnauthenticatedState when mid-auth disabled and no token returned", async () => {
            facadeChatSDK["chatConfig"] = {
                LiveWSAndLiveChatEngJoin: {
                    msdyn_authenticatedsigninoptional: "false"
                }
            } as unknown;
            facadeChatSDK["token"] = "";
            facadeChatSDK["expiration"] = 0;
            
            (handleAuthentication as jest.Mock).mockResolvedValue({ result: false, token: "", error: { message: "Auth failed" } });

            const result = await facadeChatSDK["tokenRing"]();

            expect(result.result).toBe(false);
            expect(facadeChatSDK["pendingMidAuthUnauthenticatedState"]).toBe(false);
        });

        it("should keep pendingMidAuthUnauthenticatedState when sdk is mocked (returns early)", async () => {
            facadeChatSDK["pendingMidAuthUnauthenticatedState"] = true;
            facadeChatSDK["sdkMocked"] = true; // Will return early without resetting

            await facadeChatSDK["tokenRing"]();

            // When SDK is mocked, tokenRing returns early and does not reset the flag
            expect(facadeChatSDK["pendingMidAuthUnauthenticatedState"]).toBe(true);
        });
    });

    describe("Mid-auth end-to-end scenarios", () => {
        beforeEach(() => {
            jest.clearAllMocks();
        });

        it("should handle full unauthenticated to authenticated flow", async () => {
            // Setup: Mid-auth enabled, user starts unauthenticated
            facadeChatSDK["chatConfig"] = {
                LiveWSAndLiveChatEngJoin: {
                    msdyn_authenticatedsigninoptional: "true"
                }
            } as unknown;
            facadeChatSDK["isAuthenticated"] = false;
            (facadeChatSDK["chatSDK"] as unknown).chatToken = {};

            // Step 1: Start chat unauthenticated
            const mockStartChat = jest.fn().mockResolvedValue(undefined);
            facadeChatSDK["chatSDK"].startChat = mockStartChat;
            jest.spyOn(facadeChatSDK, "tokenRing").mockResolvedValue({ result: true, message: "Token is valid" });

            await facadeChatSDK.startChat({});
            expect(mockStartChat).toHaveBeenCalled();
            expect(facadeChatSDK["isAuthenticated"]).toBe(false);

            // Step 2: User authenticates mid-conversation
            const jwt = getJWTToken();
            (facadeChatSDK["chatSDK"] as unknown).chatToken = { chatId: "test-chat-id" };
            const mockAuthenticateChat = jest.fn().mockResolvedValue(undefined);
            facadeChatSDK["chatSDK"].authenticateChat = mockAuthenticateChat;

            await facadeChatSDK.authenticateChat(jwt.token);

            // Verify authentication state
            expect(facadeChatSDK["isAuthenticated"]).toBe(true);
            expect(facadeChatSDK["token"]).toBe(jwt.token);
            expect(mockAuthenticateChat).toHaveBeenCalledWith(jwt.token, {});
        });

        it("should handle pre-auth then startChat flow", async () => {
            // Setup: Mid-auth enabled
            facadeChatSDK["chatConfig"] = {
                LiveWSAndLiveChatEngJoin: {
                    msdyn_authenticatedsigninoptional: "true"
                }
            } as unknown;
            facadeChatSDK["isAuthenticated"] = false;
            (facadeChatSDK["chatSDK"] as unknown).chatToken = {};

            // Step 1: Pre-chat authentication (before startChat)
            const jwt = getJWTToken();
            await facadeChatSDK.authenticateChat(jwt.token);

            expect(facadeChatSDK["isAuthenticated"]).toBe(true);
            expect(facadeChatSDK["token"]).toBe(jwt.token);
            expect((facadeChatSDK["chatSDK"] as unknown).authenticatedUserToken).toBe(jwt.token);

            // Step 2: Start chat with authentication
            const mockStartChat = jest.fn().mockResolvedValue(undefined);
            facadeChatSDK["chatSDK"].startChat = mockStartChat;
            jest.spyOn(facadeChatSDK, "tokenRing").mockResolvedValue({ result: true, message: "Token is valid" });

            await facadeChatSDK.startChat({});

            expect(mockStartChat).toHaveBeenCalled();
            const callArgs = mockStartChat.mock.calls[0][0];
            // deferInitialAuth should be false since user already authenticated
            expect(callArgs.deferInitialAuth ?? false).toBe(false);
        });

        it("should handle token refresh scenario", async () => {
            // Setup: Authenticated user with expiring token
            const oldJwt = getJWTToken();
            facadeChatSDK["token"] = oldJwt.token;
            facadeChatSDK["expiration"] = Math.floor(Date.now() / 1000) - 10; // Expired
            facadeChatSDK["isAuthenticated"] = true;

            // New token from handleAuthentication
            const newJwt = getJWTToken();
            (handleAuthentication as jest.Mock).mockResolvedValue({ result: true, token: newJwt.token });

            // tokenRing should get new token
            const result = await facadeChatSDK["tokenRing"]();

            expect(result.result).toBe(true);
            expect(facadeChatSDK["token"]).toBe(newJwt.token);
        });

        it("should handle reconnect scenario with hasUserAuthenticated", async () => {
            // Setup: Simulating reconnect - user was authenticated before
            const jwt = getJWTToken();
            facadeChatSDK["isAuthenticated"] = true;
            facadeChatSDK["token"] = jwt.token;
            facadeChatSDK["expiration"] = jwt.expiration;
            
            const mockStartChat = jest.fn().mockResolvedValue(undefined);
            facadeChatSDK["chatSDK"].startChat = mockStartChat;
            jest.spyOn(facadeChatSDK, "tokenRing").mockResolvedValue({ result: true, message: "Token is valid" });

            // Reconnect with liveChatContext
            const liveChatContext = { chatToken: { chatId: "reconnect-chat-id" }, requestId: "reconnect-request" };
            await facadeChatSDK.startChat({ liveChatContext } as unknown);

            expect(mockStartChat).toHaveBeenCalled();
            const callArgs = mockStartChat.mock.calls[0][0];
            // Should pass liveChatContext for reconnect
            expect(callArgs.liveChatContext).toEqual(liveChatContext);
            // deferInitialAuth should be false for authenticated reconnect
            expect(callArgs.deferInitialAuth ?? false).toBe(false);
        });
    });

    describe("Error handling and edge cases", () => {
        it("should handle concurrent authenticateChat calls gracefully", async () => {
            const jwt1 = getJWTToken();
            (facadeChatSDK["chatSDK"] as unknown).chatToken = {};

            // Call authenticateChat twice concurrently
            const promise1 = facadeChatSDK.authenticateChat(jwt1.token);
            const promise2 = facadeChatSDK.authenticateChat(jwt1.token);

            await Promise.all([promise1, promise2]);

            // Should complete without errors
            expect(facadeChatSDK["isAuthenticated"]).toBe(true);
        });

        it("should handle SDK authenticateChat timeout", async () => {
            const jwt = getJWTToken();
            (facadeChatSDK["chatSDK"] as unknown).chatToken = { chatId: "test-chat-id" };
            
            const timeoutError = new Error("Request timeout");
            const mockAuthenticateChat = jest.fn().mockRejectedValue(timeoutError);
            facadeChatSDK["chatSDK"].authenticateChat = mockAuthenticateChat;

            await expect(facadeChatSDK.authenticateChat(jwt.token)).rejects.toThrow("Request timeout");

            expect(BroadcastService.postMessage).toHaveBeenCalledWith({
                eventName: BroadcastEvent.OnWidgetError,
                payload: { errorMessage: "Request timeout" }
            });
        });

        it("should preserve authentication state after failed startChat", async () => {
            const jwt = getJWTToken();
            facadeChatSDK["isAuthenticated"] = true;
            facadeChatSDK["token"] = jwt.token;
            facadeChatSDK["expiration"] = jwt.expiration;

            const mockStartChat = jest.fn().mockRejectedValue(new Error("Network error"));
            facadeChatSDK["chatSDK"].startChat = mockStartChat;
            jest.spyOn(facadeChatSDK, "tokenRing").mockResolvedValue({ result: true, message: "Token is valid" });

            await expect(facadeChatSDK.startChat({})).rejects.toThrow("Network error");

            // Authentication state should be preserved
            expect(facadeChatSDK["isAuthenticated"]).toBe(true);
            expect(facadeChatSDK["token"]).toBe(jwt.token);
        });

        it("should handle destroy and cleanup authentication state", () => {
            const jwt = getJWTToken();
            facadeChatSDK["token"] = jwt.token;
            facadeChatSDK["expiration"] = jwt.expiration;

            facadeChatSDK.destroy();

            expect(facadeChatSDK["token"]).toBeNull();
            expect(facadeChatSDK["expiration"]).toBe(0);
        });

        it("should handle isTokenSet correctly", () => {
            expect(facadeChatSDK.isTokenSet()).toBe(false);

            facadeChatSDK["token"] = "some-token";
            expect(facadeChatSDK.isTokenSet()).toBe(true);

            facadeChatSDK["token"] = "";
            expect(facadeChatSDK.isTokenSet()).toBe(false);

            facadeChatSDK["token"] = null;
            expect(facadeChatSDK.isTokenSet()).toBe(false);
        });
    });

    describe("configureMidAuthState", () => {
        it("CASE 1: should set deferInitialAuth=true and return shouldClearReconnectParams=false for Unauth→Unauth", async () => {
            facadeChatSDK["pendingMidAuthUnauthenticatedState"] = true;
            facadeChatSDK["isAuthenticated"] = false;

            const result = await facadeChatSDK["configureMidAuthState"](false, false);

            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const sdk = facadeChatSDK["chatSDK"] as any;
            expect(sdk.deferInitialAuth).toBe(true);
            // Unauth→Unauth: liveChatContext should NOT be cleared
            expect(result.shouldClearReconnectParams).toBe(false);
        });

        it("CASE 1: should set deferInitialAuth=true and return shouldClearReconnectParams=true for Auth→Unauth", async () => {
            facadeChatSDK["pendingMidAuthUnauthenticatedState"] = true;
            facadeChatSDK["isAuthenticated"] = false;

            const result = await facadeChatSDK["configureMidAuthState"](true, true);

            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const sdk = facadeChatSDK["chatSDK"] as any;
            expect(sdk.deferInitialAuth).toBe(true);
            // Auth→Unauth: liveChatContext SHOULD be cleared
            expect(result.shouldClearReconnectParams).toBe(true);
        });

        it("CASE 2: should call handleAuthenticatedState when token is valid and not pending unauthenticated", async () => {
            const jwt = getJWTToken();
            facadeChatSDK["pendingMidAuthUnauthenticatedState"] = false;
            facadeChatSDK["token"] = jwt.token;
            facadeChatSDK["expiration"] = jwt.expiration;

            const result = await facadeChatSDK["configureMidAuthState"](false, false);

            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const sdk = facadeChatSDK["chatSDK"] as any;
            expect(sdk.authenticatedUserToken).toBe(jwt.token);
            expect(sdk.deferInitialAuth).toBe(false);
            expect(result.shouldClearReconnectParams).toBe(false);
        });

        it("should return shouldClearReconnectParams=false when no token and not pending unauthenticated", async () => {
            facadeChatSDK["pendingMidAuthUnauthenticatedState"] = false;
            facadeChatSDK["token"] = "";

            const result = await facadeChatSDK["configureMidAuthState"](false, false);

            expect(result.shouldClearReconnectParams).toBe(false);
        });
    });

    describe("handlePendingUnauthenticatedState", () => {
        it("should return true and call setMidAuthUnauthenticatedState for Auth→Unauth transition", () => {
            const spy = jest.spyOn(facadeChatSDK as any, "setMidAuthUnauthenticatedState").mockImplementation(() => {});

            const result = facadeChatSDK["handlePendingUnauthenticatedState"](true);

            expect(result).toBe(true);
            expect(spy).toHaveBeenCalled();
        });

        it("should return false and set isAuthenticated=false for Unauth→Unauth transition", () => {
            facadeChatSDK["isAuthenticated"] = true;

            const result = facadeChatSDK["handlePendingUnauthenticatedState"](false);

            expect(result).toBe(false);
            expect(facadeChatSDK["isAuthenticated"]).toBe(false);
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            expect((facadeChatSDK["chatSDK"] as any).authenticatedUserToken).toBeNull();
        });
    });

    describe("handleAuthenticatedState", () => {
        it("should set authenticatedUserToken on SDK and deferInitialAuth=false for new chat", () => {
            const jwt = getJWTToken();
            facadeChatSDK["token"] = jwt.token;

            facadeChatSDK["handleAuthenticatedState"](false, false);

            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const sdk = facadeChatSDK["chatSDK"] as any;
            expect(sdk.authenticatedUserToken).toBe(jwt.token);
            expect(sdk.deferInitialAuth).toBe(false);
        });

        it("should set deferInitialAuth=true for reconnect to unauthenticated session", () => {
            const jwt = getJWTToken();
            facadeChatSDK["token"] = jwt.token;

            facadeChatSDK["handleAuthenticatedState"](true, false);

            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const sdk = facadeChatSDK["chatSDK"] as any;
            expect(sdk.authenticatedUserToken).toBe(jwt.token);
            // Reconnect to unauth session needs migration, so defer
            expect(sdk.deferInitialAuth).toBe(true);
        });

        it("should set deferInitialAuth=false for reconnect to authenticated session", () => {
            const jwt = getJWTToken();
            facadeChatSDK["token"] = jwt.token;

            facadeChatSDK["handleAuthenticatedState"](true, true);

            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const sdk = facadeChatSDK["chatSDK"] as any;
            expect(sdk.authenticatedUserToken).toBe(jwt.token);
            // Reconnect to auth session: SDK handles auth internally
            expect(sdk.deferInitialAuth).toBe(false);
        });
    });

    describe("migrateConversationToAuthenticated", () => {
        it("should call chatSDK.authenticateChat with token and refreshChatToken=true", async () => {
            const jwt = getJWTToken();
            facadeChatSDK["token"] = jwt.token;
            const mockAuthenticateChat = jest.fn().mockResolvedValue(undefined);
            facadeChatSDK["chatSDK"].authenticateChat = mockAuthenticateChat;

            await facadeChatSDK["migrateConversationToAuthenticated"](true);

            expect(mockAuthenticateChat).toHaveBeenCalledWith(jwt.token, { refreshChatToken: true });
            expect(facadeChatSDK["isAuthenticated"]).toBe(true);
        });

        it("should not throw when authenticateChat fails (non-fatal)", async () => {
            const jwt = getJWTToken();
            facadeChatSDK["token"] = jwt.token;
            const mockAuthenticateChat = jest.fn().mockRejectedValue(new Error("Migration failed"));
            facadeChatSDK["chatSDK"].authenticateChat = mockAuthenticateChat;

            // Should not throw - migration errors are non-fatal
            await expect(facadeChatSDK["migrateConversationToAuthenticated"](true)).resolves.toBeUndefined();
        });
    });

    describe("startChat broadcast logic", () => {
        let mockStartChat: jest.Mock;

        beforeEach(() => {
            jest.clearAllMocks();
            facadeChatSDK["chatConfig"] = {
                LiveWSAndLiveChatEngJoin: {
                    msdyn_authenticatedsigninoptional: "true"
                }
            } as unknown;
            mockStartChat = jest.fn().mockResolvedValue(undefined);
            facadeChatSDK["chatSDK"].startChat = mockStartChat;
            jest.spyOn(facadeChatSDK, "tokenRing").mockResolvedValue({ result: true, message: "Token is valid" });
        });

        it("should broadcast MidConversationAuthSucceeded for new authenticated chat", async () => {
            const jwt = getJWTToken();
            facadeChatSDK["token"] = jwt.token;
            facadeChatSDK["expiration"] = jwt.expiration;
            facadeChatSDK["isAuthenticated"] = true;

            await facadeChatSDK.startChat({});

            expect(BroadcastService.postMessage).toHaveBeenCalledWith({
                eventName: BroadcastEvent.MidConversationAuthSucceeded,
                payload: {
                    isAuthenticated: true,
                    isStartChatComplete: true,
                    isReconnect: false
                }
            });
        });

        it("should broadcast MidConversationAuthReset for new unauthenticated chat", async () => {
            facadeChatSDK["pendingMidAuthUnauthenticatedState"] = true;
            facadeChatSDK["token"] = "";
            facadeChatSDK["isAuthenticated"] = false;

            await facadeChatSDK.startChat({});

            expect(BroadcastService.postMessage).toHaveBeenCalledWith({
                eventName: BroadcastEvent.MidConversationAuthReset,
                payload: {
                    isAuthenticated: false,
                    isStartChatComplete: true,
                    isReconnect: false
                }
            });
        });

        it("should skip broadcast on reconnect when auth state unchanged", async () => {
            const jwt = getJWTToken();
            facadeChatSDK["token"] = jwt.token;
            facadeChatSDK["expiration"] = jwt.expiration;
            facadeChatSDK["isAuthenticated"] = true;

            jest.clearAllMocks();
            mockStartChat = jest.fn().mockResolvedValue(undefined);
            facadeChatSDK["chatSDK"].startChat = mockStartChat;
            jest.spyOn(facadeChatSDK, "tokenRing").mockResolvedValue({ result: true, message: "Token is valid" });

            // Reconnect to authenticated session (auth state unchanged: true→true)
            await facadeChatSDK.startChat({ liveChatContext: { chatToken: {} }, wasAuthenticated: true } as any);

            // Should NOT broadcast since auth state didn't change
            expect(BroadcastService.postMessage).not.toHaveBeenCalledWith(
                expect.objectContaining({
                    eventName: BroadcastEvent.MidConversationAuthSucceeded,
                    payload: expect.objectContaining({ isStartChatComplete: true })
                })
            );
        });

        it("should trigger migrateConversationToAuthenticated for reconnect to unauthenticated session with valid token", async () => {
            const jwt = getJWTToken();
            facadeChatSDK["token"] = jwt.token;
            facadeChatSDK["expiration"] = jwt.expiration;
            facadeChatSDK["isAuthenticated"] = true;

            const mockAuthenticateChat = jest.fn().mockResolvedValue(undefined);
            facadeChatSDK["chatSDK"].authenticateChat = mockAuthenticateChat;

            // Reconnect to unauthenticated session (wasAuthenticated=false) with valid token
            await facadeChatSDK.startChat({ liveChatContext: { chatToken: {} }, wasAuthenticated: false } as any);

            // Should call authenticateChat to migrate to authenticated
            expect(mockAuthenticateChat).toHaveBeenCalledWith(jwt.token, { refreshChatToken: true });
        });

        it("should NOT trigger migrateConversationToAuthenticated for new chat", async () => {
            const jwt = getJWTToken();
            facadeChatSDK["token"] = jwt.token;
            facadeChatSDK["expiration"] = jwt.expiration;
            facadeChatSDK["isAuthenticated"] = true;

            const mockAuthenticateChat = jest.fn();
            facadeChatSDK["chatSDK"].authenticateChat = mockAuthenticateChat;

            // New chat (no liveChatContext, no reconnectId)
            await facadeChatSDK.startChat({});

            // Should NOT call authenticateChat for new chats
            expect(mockAuthenticateChat).not.toHaveBeenCalled();
        });
    });

    describe("destroy with mid-auth", () => {
        it("should reset pendingMidAuthUnauthenticatedState and deferInitialAuth when mid-auth is enabled", () => {
            facadeChatSDK["chatConfig"] = {
                LiveWSAndLiveChatEngJoin: {
                    msdyn_authenticatedsigninoptional: "true"
                }
            } as unknown;
            facadeChatSDK["pendingMidAuthUnauthenticatedState"] = true;
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            (facadeChatSDK["chatSDK"] as any).deferInitialAuth = true;

            facadeChatSDK.destroy();

            expect(facadeChatSDK["pendingMidAuthUnauthenticatedState"]).toBe(false);
            expect(facadeChatSDK["isAuthenticated"]).toBe(true);
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            expect((facadeChatSDK["chatSDK"] as any).deferInitialAuth).toBe(false);
        });

        it("should NOT reset mid-auth state when mid-auth is disabled", () => {
            facadeChatSDK["chatConfig"] = {
                LiveWSAndLiveChatEngJoin: {
                    msdyn_authenticatedsigninoptional: "false"
                }
            } as unknown;
            facadeChatSDK["pendingMidAuthUnauthenticatedState"] = true;
            facadeChatSDK["isAuthenticated"] = false;

            facadeChatSDK.destroy();

            // pendingMidAuthUnauthenticatedState should remain unchanged when mid-auth is disabled
            expect(facadeChatSDK["pendingMidAuthUnauthenticatedState"]).toBe(true);
            expect(facadeChatSDK["isAuthenticated"]).toBe(false);
        });

        it("should always clear token and expiration regardless of mid-auth", () => {
            const jwt = getJWTToken();
            facadeChatSDK["token"] = jwt.token;
            facadeChatSDK["expiration"] = jwt.expiration;

            facadeChatSDK.destroy();

            expect(facadeChatSDK["token"]).toBeNull();
            expect(facadeChatSDK["expiration"]).toBe(0);
        });
    });

    describe("clearAuthState", () => {
        it("should clear all auth-related state", () => {
            const jwt = getJWTToken();
            facadeChatSDK["token"] = jwt.token;
            facadeChatSDK["expiration"] = jwt.expiration;
            facadeChatSDK["isAuthenticated"] = true;
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            (facadeChatSDK["chatSDK"] as any).authenticatedUserToken = jwt.token;

            facadeChatSDK["clearAuthState"]();

            expect(facadeChatSDK["token"]).toBe("");
            expect(facadeChatSDK["expiration"]).toBe(0);
            expect(facadeChatSDK["isAuthenticated"]).toBe(false);
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            expect((facadeChatSDK["chatSDK"] as any).authenticatedUserToken).toBeNull();
        });
    });

});