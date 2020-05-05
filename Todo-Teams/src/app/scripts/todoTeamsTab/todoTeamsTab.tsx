import * as React from "react";
import {
    PrimaryButton,
    TeamsThemeContext,
    Panel,
    PanelBody,
    PanelHeader,
    PanelFooter,
    Surface,
    getContext,
    Input,
    SecondaryButton
} from "msteams-ui-components-react";
import TeamsBaseComponent, { ITeamsBaseComponentProps, ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import * as AuthenticationContext from "adal-angular";
import { DefaultHttpClient, TokenCredentials, ServiceClient, RequestPrepareOptions } from "@azure/ms-rest-js";
import * as query from "query-extractor";



/**
 * State for the todoTeamsTabTab React component
 */
export interface ITodoTeamsTabState extends ITeamsBaseComponentState {
    entityId?: string;
    status?: string;
    errorMessage?: string;
    showLoginButton: boolean;
    tasks: any[];
    newTask?: string;
    addButtonEnabled: boolean;
    token: string;
    tid?: string;
    teamId?: string;
    teamsTasksOnly: boolean;
    noOfItems: number;
}

/**
 * Properties for the todoTeamsTabTab React component
 */
export interface ITodoTeamsTabProps extends ITeamsBaseComponentProps {

}

/**
 * Implementation of the Todo Teams content page
 */
export class TodoTeamsTab extends TeamsBaseComponent<ITodoTeamsTabProps, ITodoTeamsTabState> {

    private authConfig: AuthenticationContext.Options;
    private authContext: AuthenticationContext;
    private resourceId: string = `https://graph.microsoft.com`;

    public constructor(props: ITodoTeamsTabProps, state: ITodoTeamsTabState) {
        super(props, state);
        this.signIn = this.signIn.bind(this);
        this.getTasks = this.getTasks.bind(this);
        this.login = this.login.bind(this);
        this.onValueChanged = this.onValueChanged.bind(this);
        this.onClickAdd = this.onClickAdd.bind(this);
    }

    public componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));
        this.setState({
            fontSize: this.pageFontSize()
        });

        if (this.inTeams()) {
            microsoftTeams.initialize();
            microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);
            microsoftTeams.getContext((context) => {
                const q = query.getAll(document.location.href);
                this.setState({
                    entityId: context.entityId,
                    tid: context.tid,
                    teamId: context.teamId,
                    teamsTasksOnly: q.teamsOnly === "true",
                    noOfItems: q.noOfItems
                });

                // Create the config
                this.authConfig = {
                    tenant: context.tid,
                    clientId: `${process.env.CLIENT_APP_ID}`,
                    endpoints: {
                        resourceId: this.resourceId
                    },
                    redirectUri: window.location.origin + "/todoTeamsTab/silentEnd.html",
                    cacheLocation: "localStorage",
                    navigateToLoginRequestUrl: false,
                };

                // Set up the auth context
                const upn = context.upn;
                if (upn) {
                    this.authConfig.extraQueryParameter = "scope=openid+profile&login_hint=" + encodeURIComponent(upn);
                } else {
                    this.authConfig.extraQueryParameter = "scope=openid+profile";
                }
                this.authContext = new AuthenticationContext(this.authConfig);

            });
        } else {
            this.setState({
                entityId: "This is not hosted in Microsoft Teams"
            });
        }
    }

    public componentDidMount() {
        if (this.inTeams()) {
            microsoftTeams.initialize();
            microsoftTeams.getContext(context => {
                this.signIn();
            });
        } else {
            this.signIn();
        }
    }

    public async getTasks(): Promise<void> {
        if (!this.state.token) {
            this.setState({
                showLoginButton: true,
                status: "No token"
            });
            return;
        }

        this.setState({
            status: "Fetching data"
        });
        const idValue = `String {390c4bde-b3f2-401b-b0a6-282eee49ad95} Name Team`;
        const extendedProp = `$filter=id eq '${idValue}'`;
        let filter = "";
        if (this.state.teamsTasksOnly) {
            const inner = `singleValueExtendedProperties/Any(ep: ep/id eq '${idValue}' and ep/value eq '${this.state.teamId as string}')`;
            filter = `&$filter=${encodeURIComponent(inner)}`;
        }
        const url = `https://graph.microsoft.com/beta/me/outlook/tasks/?$top=${this.state.noOfItems}${filter}&$expand=singleValueExtendedProperties(${encodeURI(extendedProp)})`;

        const credentials = new TokenCredentials(this.state.token);
        const client = new ServiceClient(credentials, {});
        const request: RequestPrepareOptions = {
            url,
            method: "GET"
        };
        const result = await client.sendRequest(request);
        this.setState({
            tasks: result.parsedBody.value,
            status: ""
        });

    }


    public signIn() {
        this.setState({
            status: "Logging in..."
        });

        const isCallback = this.authContext.isCallback(window.location.hash);
        if (isCallback) {
            this.authContext.handleWindowCallback();
        }

        const token = this.authContext.getCachedToken(this.resourceId);

        if (token) {
            // don't read the data in the adal renew frame
            if (!this.isInAdalRenewFrame()) {
                this.setState({ token }, () => { this.getTasks(); });
            }
        } else {
            const ac: any = this.authContext;
            ac._renewToken(this.resourceId, (errorDesc, renewdToken, error, tokenType) => {
                if (error) {
                    this.setState({
                        errorMessage: "ADAL error: " + errorDesc,
                        showLoginButton: true,
                        status: "Error"
                    });
                } else {
                    const t = this.authContext.getCachedToken(this.resourceId);
                    if (!this.isInAdalRenewFrame()) {
                        this.setState({ token: t }, () => { this.getTasks(); });
                    }
                }
            });
        }
    }

    public login() {
        if (this.inTeams()) {
            this.setState({
                status: "Requiring user input for logging in..."
            });
            microsoftTeams.authentication.authenticate({
                url: window.location.origin + "/todoTeamsTab/silentStart.html",
                width: 600,
                height: 535,
                successCallback: (result) => {
                    this.setState({
                        showLoginButton: false
                    });
                    const t = this.authContext.getCachedToken(this.resourceId);
                    this.setState({ token: t }, () => { this.getTasks(); });
                },
                failureCallback: (reason) => {
                    this.setState({
                        errorMessage: "Login failed: " + reason
                    });

                    if (reason === "CancelledByUser" || reason === "FailedToOpenWindow") {
                        this.setState({
                            status: "Login was blocked by popup blocker or canceled by user.",
                            showLoginButton: true
                        });
                    }
                }
            });
        } else {
            this.authContext.login();
        }
    }


    /**
     * The render() method to create the UI of the tab
     */
    public render() {
        const context = getContext({
            baseFontSize: this.state.fontSize,
            style: this.state.theme
        });
        const { rem, font, colors } = context;
        const { sizes, weights } = font;
        const styles = {
            header: { ...sizes.title, ...weights.semibold },
            section: { ...sizes.base, marginTop: rem(1.4), marginBottom: rem(1.4) },
            footer: { ...sizes.xsmall },
            taskTitle: { ...sizes.title2 },
            taskArea: {
                display: "flex",
                alignItems: "left",
                justifyContent: "left",
                flexDirection: "row",
                flexWrap: "wrap",
                flexFlow: "row wrap",
                alignContent: "flex-end"
            } as React.CSSProperties,
            task: {
                flex: "0 1 320px",
                height: "80px",
                ...sizes.base,
                margin: rem(1.4),
                border: rem(0.2) + " solid",
                borderRadius: rem(0.3),
                padding: rem(0.4),
                font: "inherit",
                color: colors.light.gray02,
                borderColor: colors.light.gray06,
            } as React.CSSProperties,
            teamTask: {
                flex: "0 1 320px",
                height: "80px",
                ...sizes.base,
                margin: rem(1.4),
                border: rem(0.2) + " solid",
                borderRadius: rem(0.3),
                padding: rem(0.4),
                font: "inherit",
                color: colors.light.gray02,
                borderColor: colors.light.brand00Dark,
                backgroundColor: colors.light.brand14
            } as React.CSSProperties,
            teamTaskOther: {
                flex: "0 1 320px",
                height: "80px",
                ...sizes.base,
                margin: rem(1.4),
                border: rem(0.2) + " solid",
                borderRadius: rem(0.3),
                padding: rem(0.4),
                font: "inherit",
                color: colors.light.white,
                borderColor: colors.light.brand00Dark,
                backgroundColor: colors.light.gray02
            } as React.CSSProperties,
            input: {
                padding: rem(0.5),
            }

        };
        return (
            <TeamsThemeContext.Provider value={context}>
                <Surface>
                    <Panel>
                        <PanelHeader>
                            <div style={styles.header}>Your tasks</div>
                        </PanelHeader>
                        <PanelBody>
                            <div style={styles.section}>
                                {this.state.status}
                            </div>
                            {this.state.errorMessage &&
                                <div style={styles.section}>
                                    <b>Error: </b>{this.state.errorMessage}
                                </div>
                            }
                            {this.state.showLoginButton &&
                                <div style={styles.section}>
                                    <PrimaryButton onClick={() => this.login()}>Login</PrimaryButton>
                                </div>
                            }

                            <div style={styles.section}>
                                <div style={styles.taskArea}>
                                    {this.state.tasks && this.state.tasks.map(task => {
                                        if (task.singleValueExtendedProperties &&
                                            task.singleValueExtendedProperties[0] &&
                                            (task.singleValueExtendedProperties[0].value === this.state.teamId ||
                                                task.singleValueExtendedProperties[0].value === "Test")) {
                                            return (<div style={styles.teamTask}><div style={styles.taskTitle}>{task.subject}</div></div>);
                                        } else if (task.singleValueExtendedProperties &&
                                            task.singleValueExtendedProperties[0]) {
                                            return (<div style={styles.teamTaskOther}><div style={styles.taskTitle}>{task.subject}</div></div>);
                                        } else {
                                            return (<div style={styles.task}><div style={styles.taskTitle}>{task.subject}</div></div>);

                                        }

                                    })}
                                </div>
                            </div>
                            <div style={styles.section}>
                                Add task:
                               <Input
                                    style={styles.input}
                                    placeholder="New task"
                                    value={this.state.newTask}
                                    onChange={this.onValueChanged} />
                                <PrimaryButton
                                    disabled={!this.state.addButtonEnabled}
                                    onClick={this.onClickAdd}
                                >Add task</PrimaryButton>
                            </div>

                        </PanelBody>
                        <PanelFooter>
                            <div style={styles.footer}>
                                (C) Copyright Avanade
                            </div>
                        </PanelFooter>
                    </Panel>
                </Surface>
            </TeamsThemeContext.Provider>
        );
    }
    public onValueChanged(event) {
        this.setState({
            newTask: event.target.value,
            addButtonEnabled: event.target.value.length > 0
        });
    }

    public async onClickAdd(event): Promise<void> {
        const task: any = {
            subject: this.state.newTask,
            singleValueExtendedProperties: [
                {
                    id: "String {390c4bde-b3f2-401b-b0a6-282eee49ad95} Name Team",
                    value: this.state.teamId
                }]
        };

        const credentials = new TokenCredentials(this.state.token);
        const client = new ServiceClient(credentials, {});
        const request: RequestPrepareOptions = {
            url: `https://graph.microsoft.com/beta/me/outlook/tasks/`,
            method: "POST",
            body: task
        };
        const result = await client.sendRequest(request);
        this.setState({ newTask: "" });
        this.getTasks();
    }



    public isInAdalRenewFrame() {
        return window.frameElement &&
            (window.frameElement.id.indexOf("adalRenewFrame") !== -1);
    }
}
