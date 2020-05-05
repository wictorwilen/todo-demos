import * as React from 'react';
import styles from './ToDo.module.scss';
import { IToDoProps, IToDoState, TasksResponse } from './IToDoProps';
import { Spinner } from "office-ui-fabric-react/lib/Spinner";
import { MessageBar, MessageBarType } from "office-ui-fabric-react/lib/MessageBar";
import { MSGraphClient } from '@microsoft/sp-http';
import { DocumentCard, DocumentCardTitle } from "office-ui-fabric-react/lib/DocumentCard";
import { Label } from "office-ui-fabric-react/lib/Label";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { Button, PrimaryButton } from "office-ui-fabric-react/lib/Button";
import { Customizer } from 'office-ui-fabric-react/lib/Utilities';
import { getTheme, ITheme, loadTheme } from 'office-ui-fabric-react/lib/Styling';

export default class ToDo extends React.Component<IToDoProps, IToDoState> {

  constructor(props: IToDoProps, state: IToDoState) {
    super(props, state);
    this.state = {
      loading: true,
      tasks: undefined,
      errorMessage: undefined,
      theme: "default"
    };
    this.addTask = this.addTask.bind(this);
    this.getTasks = this.getTasks.bind(this);
  }

  public componentDidMount() {
    if (this.props.context.microsoftTeams) {
      // we're actually running in Teams!
      this.props.context.microsoftTeams.getContext(context => {
        // we now have the Teams context
        this.setState({
          teamId: context.teamId,
          theme: context.theme
        });
        this.getTasks();
        this.props.context.microsoftTeams.registerOnThemeChangeHandler(theme => {
          this.setState({ theme: theme ? theme : "default" });
        });
      });
    } else {
      this.getTasks();
    }
  }


  private getTasks() {
    this.props.context.msGraphClientFactory.getClient().then((client: MSGraphClient) => {
      const filter = "$filter=id eq 'String {390c4bde-b3f2-401b-b0a6-282eee49ad95} Name Team'";
      client
        .api(`/me/outlook/tasks`)
        .top(this.props.itemCount)
        .expand(`singleValueExtendedProperties(${encodeURI(filter)})`)
        .version("beta")
        .get((error, response, rawResponse) => {
          if (error) {
            this.setState({ errorMessage: error.message });
          }
          else {
            this.setState({
              errorMessage: undefined,
              tasks: response.value,
              loading: false
            });
          }
        });
    }).catch(err => { this.setState({ errorMessage: err }); });
  }

  private addTask() {
    this.setState({
      loading: true
    });

    this.props.context.msGraphClientFactory.getClient().then((client: MSGraphClient) => {
      const task: any = {
        subject: this.state.newTaskName,
      };
      // #region Add Teams prop
      //if (this.state.teamId) {
        // If we're in a team, add the Team Id as property
        task.singleValueExtendedProperties = [
          {
            id: "String {390c4bde-b3f2-401b-b0a6-282eee49ad95} Name Team",
            value: "Test"//this.state.teamId
          }
        ];
      //}
      // #endregion

      client
        .api(`/me/outlook/tasks`)
        .version("beta")
        .post(task, (error, response, rawResponse) => {
          if (error) {
            this.setState({ errorMessage: error.message });
          }
          else {
            this.setState({
              loading: false,
              newTaskName: ''
            });
            this.getTasks();
          }
        });
    }).catch(err => { this.setState({ errorMessage: err }); });
  }

  public render(): React.ReactElement<IToDoProps> {
    console.log(this.state);
    let theme = getTheme();
    // #region Teams theme
    if (this.state.teamId) {
      theme = loadTheme({
        isInverted: this.state.theme != "default",
        palette: {
          accent: this.state.theme == 'contrast' ? 'yellow' : this.state.theme == 'default' ? '#5558AF' : '#16233A'
        },
        semanticColors: {
          bodyBackground: this.state.theme == 'default' ? "#FFFFFF" : '#16233A',
          bodyText: this.state.theme == 'default' ? "#16233A" : '#FFFFFF',
          bodySubtext: this.state.theme == 'default' ? "#16233A" : '#FFFFFF',
          menuItemText: this.state.theme == 'default' ? "#16233A" : '#FFFFFF',
          inputPlaceholderText: this.state.theme == 'default' ? "#FFFFFF" : '#16233A',
          inputText: this.state.theme == 'default' ? "#FFFFFF" : '#16233A',
        }
      });
    }
    // #endregion
    console.log(theme);

    return (
      <Customizer settings={{ theme: theme }}>
        <div className={styles.toDo} >
          {
            this.state.errorMessage &&
            <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
              {this.state.errorMessage}
            </MessageBar>
          }
          {this.state.loading &&
            <Spinner />
          }

          <div className={styles.taskContainer}>
            {this.state.tasks && this.state.tasks.map(task => {
              let cn = '';
              if (task.singleValueExtendedProperties &&
                task.singleValueExtendedProperties[0] &&
                (task.singleValueExtendedProperties[0].value == this.state.teamId ||
                  task.singleValueExtendedProperties[0].value == 'Test')) {
                cn = styles.inDaTeam;
              }
              return (<DocumentCard className={cn}>
                <DocumentCardTitle title={task.subject} />
              </DocumentCard>);
            })}
          </div>
          <div>
            <TextField label="New Task" value={this.state.newTaskName} onChanged={(v) => { this.setState({ newTaskName: v }); }} />
            <PrimaryButton label="Add" onClick={this.addTask} >Add new task</PrimaryButton>
          </div>
        </div >
      </Customizer>
    );
  }
}
