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
import { SPHttpClient } from "@microsoft/sp-http";

import * as strings from "AnnouncementDashboardWebPartStrings";
import AnnouncementDashboard from "./components/AnnouncementDashboard";
import { IAnnouncementDashboardProps } from "./components/IAnnouncementDashboardProps";

import { UI } from "./loc/il8n/ui";

export interface IAnnouncementDashboardWebPartProps {
  description: string;

  title: string;
  IsFiltering: boolean;
  Items: number;
  Layout: string;
  ListName: string;
  Language: string;
  color: string;
  sphttpclient: SPHttpClient;
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
        sphttpclient: this.context.spHttpClient,
        siteUrl: this.context.pageContext.web.absoluteUrl,
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
    this._applyPropertyPaneRtl(this.properties.Language === "AR");

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
    this._applyPropertyPaneRtl(false); // cleanup
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: any,
    newValue: any,
  ): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (propertyPath === "Language" && oldValue !== newValue) {
      // Re-render web part + refresh pane labels/options
      this._applyPropertyPaneRtl(newValue === "AR");
      this.context.propertyPane.refresh();
      this.render();
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const isAr = this.properties.Language === "AR";
    const t = isAr ? UI.AR : UI.EN;
    return {
      pages: [
        {
          groups: [
            {
              groupName: "Announcement Dashboard Settings",
              groupFields: [
                PropertyPaneTextField("title", {
                  label: t.WebPartTitle,
                }),

                PropertyPaneToggle("IsFiltering", {
                  label: t.IsFiltering,
                  onText: t.FilteringEnabled,
                  offText: t.FilteringDisabled,
                }),

                PropertyPaneSlider("Items", {
                  label: t.ItemsLabel,
                  min: 1,
                  max: 10,
                  step: 1,
                  showValue: true,
                }),
                PropertyPaneChoiceGroup("Layout", {
                  label: t.LayoutLabel,
                  options: [
                    { key: "Card", text: t.Card },
                    { key: "Compact", text: t.Compact },
                    { key: "Table", text: t.Table },
                  ],
                }),
                PropertyPaneDropdown("ListName", {
                  label: t.ListnameLabel,
                  selectedKey: "Announcement",
                  options: [{ key: "Announcement", text: "Announcement" }],
                }),

                PropertyPaneChoiceGroup("Language", {
                  label: t.LanguageLabel,
                  options: [
                    { key: "EN", text: "English" },
                    { key: "AR", text: "Arabic" },
                  ],
                }),

                PropertyFieldColorPicker("color", {
                  label: t.ColorLabel,
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
  private _rtlStyleTagId = "ad-propane-rtl";

  private _applyPropertyPaneRtl(isAr: boolean): void {
    const id = this._rtlStyleTagId;
    const existing = document.getElementById(id);

    if (!isAr) {
      existing?.remove();
      return;
    }

    const css = `
/* ========= Base RTL for property pane ========= */
.spPropertyPaneContainer {
  direction: rtl !important;
}

/* Labels / text */
.spPropertyPaneContainer label,
.spPropertyPaneContainer .ms-Label
{
  direction: rtl !important;
  text-align: right !important;
}
.spPropertyPaneContainer .fui-Label {
text-align: right !important;
display:flex !important;
justify-content: flex-start !important;
}
/* Text inputs */
.spPropertyPaneContainer input[type="text"],
.spPropertyPaneContainer textarea {
  direction: rtl !important;
  text-align: right !important;
}

/* ========= Slider: keep mechanics LTR ========= */
.spPropertyPaneContainer .ms-Slider {
  direction: ltr !important;
}
.spPropertyPaneContainer .ms-Slider-container {
  display: flex !important;
  flex-direction: row-reverse !important;
  align-items: center !important;
}
.spPropertyPaneContainer .ms-Slider-value {
  margin-right: 12px !important;
  margin-left: 0 !important;
  text-align: right !important;
}

/* ========= Dropdown ========= */
.spPropertyPaneContainer .ms-Dropdown {
  direction: rtl !important;
}
.spPropertyPaneContainer .ms-Dropdown-title {
  text-align: right !important;
}
.spPropertyPaneContainer .ms-Dropdown-caretDownWrapper {
  left: auto !important;
  right: 8px !important;
}

/* ============================================================
   FIX 1: Switch (تفعيل التصفية)
   - Switch pill stays on right in the pane
   - Thumb movement behaves like EN (OFF left, ON right)
   - Whole switch row clickable (indicator + text)
   ============================================================ */
.spPropertyPaneContainer
  [data-automation-id="propertyPaneGroupField"]:has(input[role="switch"][aria-label="تفعيل التصفية"])
  > div {
  direction: rtl !important;
  text-align: right !important;
}

/* Layout: keep pill on the right */
.spPropertyPaneContainer
  [data-automation-id="propertyPaneGroupField"]:has(input[role="switch"][aria-label="تفعيل التصفية"])
  .fui-Switch {
  position: relative !important;
  display: flex !important;
  flex-direction: row-reverse !important; /* pill right */
  align-items: center !important;
  gap: 10px !important;
  width: 100% !important;

  /* IMPORTANT: keep switch mechanics LTR so thumb moves correctly */
  direction: ltr !important;
}

/* Make entire switch row clickable */
.spPropertyPaneContainer
  [data-automation-id="propertyPaneGroupField"]:has(input[role="switch"][aria-label="تفعيل التصفية"])
  .fui-Switch__input {
  position: absolute !important;
  inset: 0 !important;     
  width: 100% !important;
  height: 100% !important;
  opacity: 0 !important;
  cursor: pointer !important;
  z-index: 10 !important;
}

/* Keep label RTL */
.spPropertyPaneContainer
  [data-automation-id="propertyPaneGroupField"]:has(input[role="switch"][aria-label="تفعيل التصفية"])
  .fui-Switch__label {
  direction: rtl !important;
  text-align: right !important;
}

/* ============================================================
   FIX 2: Radio / ChoiceGroup
   - Make entire option row clickable (circle + text + whitespace)
   - Keep circle on the right for Arabic
   ============================================================ */
.spPropertyPaneContainer .fui-Radio {
  position: relative !important;
  display: flex !important;
  flex-direction: row-reverse !important;
  align-items: center !important;
  width: 100% !important;
  cursor: pointer !important;
}

/* Cover whole option with the input => everything clickable */
.spPropertyPaneContainer .fui-Radio__input {
  position: absolute !important;
  inset: 0 !important;    
  width: 100% !important;
  height: 100% !important;
  opacity: 0 !important;
  cursor: pointer !important;
  z-index: 10 !important;
}

.spPropertyPaneContainer .fui-Radio__indicator,
.spPropertyPaneContainer .fui-Radio__label {
  position: relative !important;
  z-index: 2 !important;
  cursor: pointer !important;
}

.spPropertyPaneContainer .fui-Radio__label {
  direction: rtl !important;
  text-align: right !important;
  flex: 1 !important;
}


`;

    if (!existing) {
      const style = document.createElement("style");
      style.id = id;
      style.textContent = css;
      document.head.appendChild(style);
    } else {
      existing.textContent = css;
    }
  }
}
