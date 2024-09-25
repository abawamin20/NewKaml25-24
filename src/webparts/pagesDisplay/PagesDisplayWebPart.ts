import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane"; // Import PropertyPaneTextField
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

import * as strings from "PagesDisplayWebPartStrings";
import PagesDisplay from "./components/PagesDisplay";
import { IPagesDisplayProps } from "./components/IPagesDisplayProps";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

export interface IPagesDisplayWebPartProps {
  selectedView: string;
  feedbackLink: string; // Add this line
}

export interface IView {
  Id: string;
  Title: string;
}

export default class PagesDisplayWebPart extends BaseClientSideWebPart<IPagesDisplayWebPartProps> {
  private viewOptions: IView[] = [];

  public render(): void {
    const element: React.ReactElement<IPagesDisplayProps> = React.createElement(
      PagesDisplay,
      {
        context: this.context,
        selectedViewId: this.properties.selectedView,
        feedbackPageUrl: this.properties.feedbackLink, // Pass the feedback link to the child component
      }
    );

    ReactDom.render(element, this.domElement);
  }

  public async getViews(): Promise<void> {
    const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Site Pages')/views`;
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Failed to fetch views. Error code: ${response.status}`);
    }

    const data = await response.json();
    this.viewOptions = [
      {
        Id: "",
        Title: "Select View",
      },
      ...data.value.map((view: any) => ({
        Id: view.Id,
        Title: view.Title,
      })),
    ];
  }

  protected onInit(): Promise<void> {
    return this.getViews();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty(
        "--bodyText",
        semanticColors.bodyText || null
      );
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty(
        "--linkHovered",
        semanticColors.linkHovered || null
      );
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Configure your pages display",
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneDropdown("selectedView", {
                  label: "Select View",
                  options: this.viewOptions.map((view) => ({
                    key: view.Id,
                    text: view.Title,
                  })),
                }),
                PropertyPaneTextField("feedbackLink", {
                  // Add this block
                  label: "Feedback Link",
                  description: "Enter the URL for the feedback link",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
