import { PreventIframe } from "express-msteams-host";

/**
 * Used as place holder for the decorators
 */
@PreventIframe("/todoTeamsTab/index.html")
@PreventIframe("/todoTeamsTab/config.html")
@PreventIframe("/todoTeamsTab/remove.html")
export class TodoTeamsTab {
    
}