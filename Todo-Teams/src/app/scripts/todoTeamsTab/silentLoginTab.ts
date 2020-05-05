import * as microsoftTeams from "@microsoft/teams-js";
import * as AuthenticationContext from "adal-angular";

export class SilentLoginTab {

    public static Start() {
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            const config: AuthenticationContext.Options = {
                tenant: context.tid,
                clientId: `${process.env.CLIENT_APP_ID}`,
                endpoints: {
                    resourceId: this.resourceId
                },
                redirectUri: window.location.origin + "/todoTeamsTab/silentEnd.html",
                cacheLocation: "localStorage",
                navigateToLoginRequestUrl: false,
            };
            if (context.upn) {
                config.extraQueryParameter = "scope=openid+profile&login_hint=" + encodeURIComponent(context.upn);
            } else {
                config.extraQueryParameter = "scope=openid+profile";
            }
            const authContext = new AuthenticationContext(config);
            authContext.login();
        });
    }

    public static End() {
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            const config: AuthenticationContext.Options = {
                tenant: context.tid,
                clientId: `${process.env.CLIENT_APP_ID}`,
                endpoints: {
                    resourceId: this.resourceId
                },
                redirectUri: window.location.origin + "/todoTeamsTab/silentEnd.html",
                cacheLocation: "localStorage",
                navigateToLoginRequestUrl: false,
            };
            const authContext = new AuthenticationContext(config);
            if (authContext.isCallback(window.location.hash)) {
                authContext.handleWindowCallback(window.location.hash);
                if (authContext.getCachedUser()) {
                    authContext.acquireToken(this.resourceId, (errorDesc, token, error) => {
                        microsoftTeams.authentication.notifySuccess(token as string);
                    });

                } else {
                    microsoftTeams.authentication.notifyFailure(authContext.getLoginError());
                }
            }
        });
    }

    protected static resourceId: string = `https://graph.microsoft.com`;

}
