import * as React from "react";
import styles from "./AnnouncementDashboard.module.scss";
import type { IAnnouncementDashboardProps } from "./IAnnouncementDashboardProps";
import { escape } from "@microsoft/sp-lodash-subset";

import $ from "jquery";
import moment from "moment";
import { DetailsList, DetailsListLayoutMode } from "@fluentui/react";
import { UI } from "../loc/il8n/ui";

import * as strings from "AnnouncementDashboardWebPartStrings";
const formatDate = (dateString: string): string => {
  return moment(dateString).format("M/DD/YYYY");
};
export interface IDetailListItem {
  Title_ar: string;
  Title_en: string;
  Description_ar: string;
  Description_en: string;
  // category: number;
  Priority: number;
  DueDate: string;
}
export interface IDetailListState {
  items: IDetailListItem[];
}

export interface IAnnouncementDashboardState {
  listItems: [
    {
      Title_ar: string;
      Title_en: string;
      Description_ar: string;
      Description_en: string;
      category?: {
        Title: string;
        name_en?: string;
        name_ar?: string;
      };
      Priority: number;
      DueDate: string;
      AssignedTo?: {
        Title: string;
        EMail?: string;
      };
    },
  ];
}

export default class AnnouncementDashboard extends React.Component<
  IAnnouncementDashboardProps,
  IAnnouncementDashboardState
> {
  static siteUrl: string = "";
  public constructor(
    props: IAnnouncementDashboardProps,
    state: IAnnouncementDashboardState,
  ) {
    super(props);
    this.state = {
      listItems: [
        {
          Title_ar: "",
          Title_en: "",
          Description_ar: "",
          Description_en: "",
          category: { Title: "", name_en: "", name_ar: "" },
          Priority: 0,
          DueDate: "",
          AssignedTo: { Title: "", EMail: "" },
        },
      ],
    };
    AnnouncementDashboard.siteUrl = this.props.websiteUrl;

    this.updateState = this.updateState.bind(this);
    console.log("stage title from constructor");
  }
  public componentWillMount() {
    console.log("comp will mount has been called");
  }
  public componentDidMount() {
    this.fetchAnnouncements();
  }

  public componentDidUpdate(prevProps: IAnnouncementDashboardProps): void {
    // Re-fetch data if Items, Language, or ListName changes
    if (
      prevProps.Items !== this.props.Items ||
      prevProps.Language !== this.props.Language ||
      prevProps.ListName !== this.props.ListName ||
      prevProps.Layout !== this.props.Layout
    ) {
      this.fetchAnnouncements();
    }
  }
  private fetchAnnouncements(): void {
    const apiUrl =
      `${AnnouncementDashboard.siteUrl}/sites/DevelopmentDemos/_api/web/lists/getbytitle('Announcement')/items` +
      `?$top=${this.props.Items}` +
      `&$select=Id,Title_en,Title_ar,Description_en,Description_ar,Priority,DueDate,` +
      `category/Title,category/name_en,category/name_ar,` +
      `AssignedTo/Title,AssignedTo/EMail` +
      `&$expand=category,AssignedTo`;
    console.log("api url", apiUrl);
    $.ajax({
      url: apiUrl,

      type: "GET",
      headers: {
        Accept: "application/json;odata=verbose",
      },
      success: (resultData) => {
        this.setState({
          listItems: resultData.d.results,
        });
      },
      error: (jqXHR, textStatus, errorThrown) => {
        console.error("AJAX ERROR");
        console.error("URL attempted:", apiUrl);
        console.error("Status:", textStatus);
        console.error("Error thrown:", errorThrown);
        console.error("HTTP status:", jqXHR.status);
        console.error("Response text:", jqXHR.responseText);
      },
    });
  }

  public updateState() {
    this.setState({});
  }
  public render(): React.ReactElement<IAnnouncementDashboardProps> {
    const isAr = this.props.Language === "AR";
    const t = isAr ? UI.AR : UI.EN;
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = this.props;

    return (
      // <section
      //   className={`${styles.announcementDashboard} ${hasTeamsContext ? styles.teams : ""}`}
      // >
      //   <div className={styles.welcome}>
      //     <img
      //       alt=""
      //       src={
      //         isDarkTheme
      //           ? require("../assets/welcome-dark.png")
      //           : require("../assets/welcome-light.png")
      //       }
      //       className={styles.welcomeImage}
      //     />
      //     <h2>Well done, {escape(userDisplayName)}!</h2>
      //     <div>{environmentMessage}</div>
      //     <div>
      //       Web part property value: <strong>{escape(description)}</strong>
      //     </div>
      //   </div>
      //   <div>
      //     <h3>Welcome to SharePoint Framework using react!</h3>
      //     <p>
      //       The SharePoint Framework (SPFx) is a extensibility model for
      //       Microsoft Viva, Microsoft Teams and SharePoint. It&#39;s the easiest
      //       way to extend Microsoft 365 with automatic Single Sign On, automatic
      //       hosting and industry standard tooling.
      //     </p>
      //     <h4>Learn more about SPFx development:</h4>
      //     <ul className={styles.links}>
      //       <li>
      //         <a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">
      //           SharePoint Framework Overview
      //         </a>
      //       </li>
      //       <li>
      //         <a
      //           href="https://aka.ms/spfx-yeoman-graph"
      //           target="_blank"
      //           rel="noreferrer"
      //         >
      //           Use Microsoft Graph in your solution
      //         </a>
      //       </li>
      //       <li>
      //         <a
      //           href="https://aka.ms/spfx-yeoman-teams"
      //           target="_blank"
      //           rel="noreferrer"
      //         >
      //           Build for Microsoft Teams using SharePoint Framework
      //         </a>
      //       </li>
      //       <li>
      //         <a
      //           href="https://aka.ms/spfx-yeoman-viva"
      //           target="_blank"
      //           rel="noreferrer"
      //         >
      //           Build for Microsoft Viva Connections using SharePoint Framework
      //         </a>
      //       </li>
      //       <li>
      //         <a
      //           href="https://aka.ms/spfx-yeoman-store"
      //           target="_blank"
      //           rel="noreferrer"
      //         >
      //           Publish SharePoint Framework applications to the marketplace
      //         </a>
      //       </li>
      //       <li>
      //         <a
      //           href="https://aka.ms/spfx-yeoman-api"
      //           target="_blank"
      //           rel="noreferrer"
      //         >
      //           SharePoint Framework API reference
      //         </a>
      //       </li>
      //       <li>
      //         <a href="https://aka.ms/m365pnp" target="_blank" rel="noreferrer">
      //           Microsoft 365 Developer Community
      //         </a>
      //       </li>
      //     </ul>
      //   </div>
      // </section>
      <div
        dir={this.props.Language === "AR" ? "rtl" : "ltr"}
        className={styles.announcementDashboard}
      >
        <table className={styles.table}>
          <thead style={{ backgroundColor: this.props.color }}>
            <tr>
              <th className={styles.th}>{t.HeaderTitle}</th>
              <th className={styles.th}>{t.HeaderDescription}</th>
              <th className={styles.th}>{t.HeaderCategory}</th>
              <th className={styles.th}>{t.HeaderPriority}</th>
              <th className={styles.th}>{t.HeaderDueDate}</th>
              <th className={styles.th}>{t.HeaderAssignedTo}</th>
            </tr>
          </thead>
          <tbody>
            {this.state.listItems.map((listitem, listitemkey) => {
              const isArabic = this.props.Language === "AR";
              return (
                <tr className={styles.tr} key={listitemkey}>
                  <td className={styles.td}>
                    {isArabic ? listitem.Title_ar : listitem.Title_en}
                  </td>
                  <td className={styles.td}>
                    {isArabic
                      ? listitem.Description_ar
                      : listitem.Description_en}
                  </td>
                  <td className={styles.td}>
                    {isArabic
                      ? listitem.category?.name_ar
                      : listitem.category?.name_en || "Uncategorized"}
                  </td>
                  <td className={styles.td}>{listitem.Priority}</td>
                  <td className={styles.td}>{formatDate(listitem.DueDate)}</td>
                  <td className={styles.td}>
                    {listitem.AssignedTo?.Title || "Unassigned"}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>

        <div className="propertypanedisplay">
          <h3>Web Part Title: {this.props.title}</h3>
          <h3>Is Filtering Enabled: {this.props.IsFiltering ? "Yes" : "No"}</h3>
          <h3>Number of items to display: {this.props.Items}</h3>
          <h3>Layout Style: {this.props.Layout}</h3>
          <h3>Selected List: {this.props.ListName}</h3>
          <h3>Selected Language: {this.props.Language}</h3>
          <h3>Selected Background Color: {this.props.color}</h3>
        </div>
      </div>
    );
  }

  public componentWillUnmount() {
    console.log("component will unmount has been called");
  }
}
