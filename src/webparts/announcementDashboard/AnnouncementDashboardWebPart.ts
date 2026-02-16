import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
} from "@microsoft/sp-property-pane";
import {
  PropertyFieldColorPicker,
  PropertyFieldColorPickerStyle,
} from "@pnp/spfx-property-controls/lib/propertyFields/colorPicker";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

import * as strings from "AnnouncementDashboardWebPartStrings";
import AnnouncementDashboard from "./components/AnnouncementDashboard";
import { IAnnouncementDashboardProps } from "./components/IAnnouncementDashboardProps";

export interface IAnnouncementDashboardWebPartProps {
  description: string;

  title: string;
  IsFiltering: boolean;
  Items: number;
  Layout: string;
  ListName: string;
  Language: string;
  color: string;
}

export default class AnnouncementDashboardWebPart extends BaseClientSideWebPart<IAnnouncementDashboardWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";

  protected get disableReactivePropertyChanges(): boolean {
    return false;
  }

  public render(): void {
    const element: React.ReactElement<IAnnouncementDashboardProps> =
      React.createElement(AnnouncementDashboard, {
        IsFiltering: this.properties.IsFiltering,
        title: this.properties.title,
        description: this.properties.description,
        Items: this.properties.Items,
        Layout: this.properties.Layout,
        ListName: this.properties.ListName,
        Language: this.properties.Language,
        color: this.properties.color,
        context: this.context,

        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        websiteUrl: this.context.pageContext.web.absoluteUrl,
        key: this.properties.Language,
      });

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then((message) => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app
        .getContext()
        .then((context) => {
          let environmentMessage: string = "";
          switch (context.app.host.name) {
            case "Office": // running in Office
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOffice
                : strings.AppOfficeEnvironment;
              break;
            case "Outlook": // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOutlook
                : strings.AppOutlookEnvironment;
              break;
            case "Teams": // running in Teams
            case "TeamsModern":
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentTeams
                : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(
      this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentSharePoint
        : strings.AppSharePointEnvironment,
    );
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty(
        "--bodyText",
        semanticColors.bodyText || null,
      );
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty(
        "--linkHovered",
        semanticColors.linkHovered || null,
      );
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  // protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
  //   return {
  //     pages: [
  //       {
  //         header: {
  //           description: strings.PropertyPaneDescription,
  //         },
  //         groups: [
  //           {
  //             groupName: strings.BasicGroupName,
  //             groupFields: [
  //               PropertyPaneTextField("description", {
  //                 label: strings.DescriptionFieldLabel,
  //               }),
  //               PropertyPaneTextField("title", {
  //                 label: "Web Part Title",
  //               }),
  //             ],
  //           },
  //         ],
  //       },
  //     ],
  //   };
  // }
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: "Announcement Dashboard Settings",
              groupFields: [
                PropertyPaneTextField("title", {
                  label: "Web Part Title",
                }),

                PropertyPaneToggle("IsFiltering", {
                  label: "Enable Filtering",
                  onText: "Filtering enabled",
                  offText: "Filtering disabled",
                }),

                PropertyPaneSlider("Items", {
                  label: "Number of items to display",
                  min: 1,
                  max: 10,
                  step: 1,
                  showValue: true,
                  value: 5,
                }),
                PropertyPaneChoiceGroup("Layout", {
                  label: "Select Layout style",
                  options: [
                    { key: "Card", text: "Card" },
                    { key: "Compact", text: "Compact" },
                    { key: "Table", text: "Table" },
                  ],
                }),
                PropertyPaneDropdown("ListName", {
                  label: "Select the Announcement List",
                  selectedKey: "Operations",
                  options: [
                    { key: "HR", text: "HR" },
                    { key: "Finance", text: "Finance" },
                    { key: "Operations", text: "Operations" },
                  ],
                }),

                PropertyPaneChoiceGroup("Language", {
                  label: "Select Language",
                  options: [
                    { key: "EN", text: "English" },
                    { key: "AR", text: "Arabic" },
                  ],
                }),

                PropertyFieldColorPicker("color", {
                  label: "Select Header Background Color",
                  selectedColor: this.properties.color,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  debounce: 1000,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Inline,
                  iconName: "Precipitation",
                  key: "colorFieldId",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
