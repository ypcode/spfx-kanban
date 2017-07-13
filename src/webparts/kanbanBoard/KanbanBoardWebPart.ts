import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import { SPComponentLoader } from '@microsoft/sp-loader';

import styles from './KanbanBoard.module.scss';
import * as strings from 'kanbanBoardStrings';
import { IKanbanBoardWebPartProps } from './IKanbanBoardWebPartProps';

import pnp from "sp-pnp-js";

import * as $ from "jquery";
require("jquery-ui/ui/widgets/draggable");
require("jquery-ui/ui/widgets/droppable");

import { ITask } from "./models/ITask";
import { IListInfo } from "./models/IListInfo";
import { IFieldInfo } from "./models/IFieldInfo";

export const LAYOUT_MAX_COLUMNS = 12;

export default class KanbanBoardWebPart extends BaseClientSideWebPart<IKanbanBoardWebPartProps> {

  private statuses: string[] = [];
  private tasks: ITask[] = [];
  private availableLists: IListInfo[] = [];


  constructor() {
    super();
    SPComponentLoader.loadCss('https://code.jquery.com/ui/1.12.1/themes/base/jquery-ui.min.css');
  }

  public onInit(): Promise<any> {
    return super.onInit().then(_ => {
      // Configure PnP Js for working seamlessly with SPFx
      pnp.setup({
        spfxContext: this.context
      });
    })
      // Load the available lists in the current site
      .then(() => this._loadAvailableLists())
      .then((lists: IListInfo[]) => this.availableLists = lists);
  }

  private _loadAvailableLists(): Promise<IListInfo[]> {
    return pnp.sp.web.lists.expand("Fields").select("Id", "Title", "Fields/Title", "Fields/InternalName", "Fields/TypeAsString").get();
  }

  private _loadTasks(): Promise<ITask[]> {
    return pnp.sp.web.lists.getById(this.properties.tasksListId).items.select("Id", "Title", this.properties.statusFieldName).get()
      .then((results: ITask[]) => results && results.map(t => ({
        Id: t.Id,
        Title: t.Title,
        Status: t[this.properties.statusFieldName]
      })));
  }

  private _loadStatuses(): Promise<string[]> {
    console.log("Load statuses...");
    return pnp.sp.web.lists.getById(this.properties.tasksListId).fields.getByInternalNameOrTitle(this.properties.statusFieldName).get()
      .then((fieldInfo: IFieldInfo) => fieldInfo.Choices || []);
  }

  public render(): void {
    // Only if properly configured
    if (this.properties.statusFieldName && this.properties.tasksListId) {
      // Load the data
      this._loadStatuses()
        .then((statuses: string[]) => this.statuses = statuses)
        .then(() => this._loadTasks())
        .then((tasks: ITask[]) => this.tasks = tasks)
        // And then render
        .then(() => this._renderBoard())
        .catch(error => {
          console.log(error);
          console.log("An error occured while loading data...");
        });
    } else {
      this.domElement.innerHTML = "<div>Please configure the WebPart</div>";
    }
  }

  private _getColumnSizeClassName(columnsCount: number): string {
    if (columnsCount < 1) {
      console.log("Invalid number of columns");
      return "";
    }

    if (columnsCount > LAYOUT_MAX_COLUMNS) {
      console.log("Too many columns for responsive UI");
      return "";
    }

    let columnSize = Math.floor(LAYOUT_MAX_COLUMNS / columnsCount);

    console.log("Column size =" + columnSize);
    return "ms-u-sm" + (columnSize || 1);
  }

  /**
   * Generates and inject the HTML of the Kanban Board
   */
  private _renderBoard(): void {
    let columnSizeClass = this._getColumnSizeClassName(this.statuses.length);

    // The begininning of the WebPart
    let html = `
      <div class="${styles.kanbanBoard}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ${styles.row}">`;

    // For each status
    this.statuses.forEach(status => {
      // Append a new Office UI Fabric column with the appropriate with to the row
      html += `<div class="${styles.kanbanColumn} ms-Grid-col ${columnSizeClass}" data-status="${status}">
                  <h3 class="ms-fontColor-themePrimary">${status}</h3>`;
      // Get all the tasks in the current status
      let currentGroupTasks = this.tasks.filter(t => t.Status == status);
      // Add a new tile for each task in the current column
      currentGroupTasks.forEach(task => {
        html += `<div class="${styles.task}" data-taskid="${task.Id}">
                        <div class="ms-fontSize-xl">${task.Title}</div>
                      </div>`;
      });
      // Close the column element
      html += `</div>`;
    });

    // Ends the WebPart HTML
    html += `</div>
        </div>
      </div>`;

    // Apply the generated HTML to the WebPart DOM element
    this.domElement.innerHTML = html;

    console.log("Kanban columns found : " + $(".kanban-column").length);
    // Make the kanbanColumn elements droppable areas
    let webpart = this;
    $(`.${styles.kanbanColumn}`).droppable({
      tolerance: "intersect",
      accept: `.${styles.task}`,
      activeClass: "ui-state-default",
      hoverClass: "ui-state-hover",
      drop: (event, ui) => {
        // Here the code to execute whenever an element is dropped into a column
        let taskItem = $(ui.draggable);
        let source = taskItem.parent();
        let previousStatus = source.data("status");
        let taskId = taskItem.data("taskid");
        let target = $(event.target);
        let newStatus = target.data("status");
        taskItem.appendTo(target);

        // If the status has changed, apply the changes
        if (previousStatus != newStatus) {
          webpart.changeTaskStatus(taskId, newStatus);
        }
      }
    });

    console.log("Task items found : " + $(".task").length);
    // Make the task items draggable
    $(`.${styles.task}`).draggable({
      classes: {
        "ui-draggable-dragging": styles.dragging
      },
      opacity: 0.7,
      helper: "clone",
      cursor: "move",
      revert: "invalid"
    });
  }

  private changeTaskStatus(taskId: number, newStatus: string) {
    // Set the value for the configured "status" field
    let fieldsToUpdate = {};
    fieldsToUpdate[this.properties.statusFieldName] = newStatus;

    // Update the property on the list item
    pnp.sp.web.lists.getById(this.properties.tasksListId).items.getById(taskId).update(fieldsToUpdate);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private _getAvailableFieldsFromCurrentList(): IFieldInfo[] {
    if (!this.properties.tasksListId)
      return [];

    let filteredListInfo = this.availableLists.filter(l => l.Id == this.properties.tasksListId);
    if (filteredListInfo.length != 1)
      return [];

    return filteredListInfo[0].Fields.filter(f => f.TypeAsString == "Choice");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Settings"
          },
          groups: [
            {
              groupName: "Tasks source configuration",
              groupFields: [
                PropertyPaneDropdown('tasksListId', {
                  label: "Source Task list",
                  options: this.availableLists.map(l => ({
                    key: l.Id,
                    text: l.Title
                  }))
                }),
                PropertyPaneDropdown('statusFieldName', {
                  label: "Status field Internal name",
                  options: this._getAvailableFieldsFromCurrentList().map(f => ({
                    key: f.InternalName,
                    text: f.Title
                  }))
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
