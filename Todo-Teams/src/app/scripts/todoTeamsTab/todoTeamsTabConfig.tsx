import * as React from "react";
import {
    PrimaryButton,
    Panel,
    PanelBody,
    PanelHeader,
    PanelFooter,
    Input,
    Surface,
    getContext,
    Toggle,
    TeamsThemeContext
} from "msteams-ui-components-react";
import TeamsBaseComponent, { ITeamsBaseComponentProps, ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import * as query from "query-extractor";

export interface ITodoTeamsTabConfigState extends ITeamsBaseComponentState {
    noOfItems: string;
    teamsOnly: boolean;
    entityId: string;
}

export interface ITodoTeamsTabConfigProps extends ITeamsBaseComponentProps {

}

/**
 * Implementation of Todo Teams configuration page
 */
export class TodoTeamsTabConfig extends TeamsBaseComponent<ITodoTeamsTabConfigProps, ITodoTeamsTabConfigState> {

    public componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));
        this.setState({
            fontSize: this.pageFontSize()
        });

        if (this.inTeams()) {
            microsoftTeams.initialize();

            microsoftTeams.getContext((context: microsoftTeams.Context) => {
                microsoftTeams.settings.getSettings((settings) => {
                    if (settings.contentUrl === undefined) {
                        this.setState({
                            entityId: context.entityId,
                        });
                    } else {
                        const q = query.getAll(settings.contentUrl);
                        this.setState({
                            noOfItems: q.noOfItems,
                            teamsOnly: q.teamsOnly === "true",
                            entityId: context.entityId,
                        });

                    }
                    this.setValidityState(true);

                });
            });

            microsoftTeams.settings.registerOnSaveHandler((saveEvent: microsoftTeams.settings.SaveEvent) => {
                // Calculate host dynamically to enable local debugging
                const host = "https://" + window.location.host;
                microsoftTeams.settings.setSettings({
                    contentUrl: host + `/todoTeamsTab/?noOfItems=${this.state.noOfItems}&teamsOnly=${this.state.teamsOnly}`,
                    suggestedDisplayName: "Todo Teams",
                    removeUrl: host + "/todoTeamsTab/remove.html",
                    entityId: this.state.entityId
                });
                saveEvent.notifySuccess();
            });
        } else {
        }
    }

    public render() {
        const context = getContext({
            baseFontSize: this.state.fontSize,
            style: this.state.theme
        });
        const { rem, font } = context;
        const { sizes, weights } = font;
        const styles = {
            header: { ...sizes.title, ...weights.semibold },
            section: { ...sizes.base, marginTop: rem(1.4), marginBottom: rem(1.4) },
            footer: { ...sizes.xsmall }
        };
        return (
            <TeamsThemeContext.Provider value={context}>
                <Surface>
                    <Panel>
                        <PanelHeader>
                            <div style={styles.header}>Configure your tab</div>
                        </PanelHeader>
                        <PanelBody>
                            <div style={styles.section}>
                                <Input
                                    autoFocus
                                    placeholder="Number of tasks"
                                    label="Number of tasks to show"
                                    errorLabel={!this.state.noOfItems ? "This value is required" : undefined}
                                    value={this.state.noOfItems}
                                    onChange={(e) => {
                                        this.setState({
                                            noOfItems: e.target.value
                                        });
                                    }}
                                    required />
                            </div>
                            <div style={styles.section}>
                                Only show tasks for this Team
                                <Toggle
                                    checked={this.state.teamsOnly}
                                    onToggle={(e) => {
                                        this.setState({
                                            teamsOnly: e
                                          });
                                    }}
                                />
                            </div>

                        </PanelBody>
                        <PanelFooter>
                        </PanelFooter>
                    </Panel>
                </Surface>
            </TeamsThemeContext.Provider>
        );
    }
}
