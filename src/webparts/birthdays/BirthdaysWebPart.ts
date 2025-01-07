import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'BirthdaysWebPartStrings';
import Birthdays from './components/Birthdays/Birthdays';
import { IBirthdaysProps } from './components/Birthdays/IBirthdaysProps';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy, PropertyFieldNumber } from '@pnp/spfx-property-controls';
import getSP from '../PnPjsConfig';
import { SPFI } from "@pnp/sp";

const { solution } = require("../../../config/package-solution.json");

export interface IBirthdaysWebPartProps {
  description: string;
  BirthdayListId: string;
  BirthdaysRange: string;
  BackgroundImage: string;
  pageItemsNumber: number;

  SpotlightId: string;
  SpotlightInterval: number;
}

export default class BirthdaysWebPart extends BaseClientSideWebPart<IBirthdaysWebPartProps> {

  sp: SPFI

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IBirthdaysProps> = React.createElement(
      Birthdays,
      {
        description: this.properties.description,
        sp: this.sp,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        BirthdayListId: this.properties.BirthdayListId,
        BirthdaysRange: this.properties.BirthdaysRange,
        context: this.context,
        BackgroundImage: this.properties.BackgroundImage,
        pageItemsNumber: this.properties.pageItemsNumber,

        SpotlightId: this.properties.SpotlightId,
        SpotlightInterval: this.properties.SpotlightInterval,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this.sp = getSP(this.context)
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse(solution.version);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyFieldListPicker("NewsListId", {
                  label: "Select News list",
                  selectedList: this.properties.BirthdayListId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                  key: "listPickerFieldId",
                }),
                PropertyPaneTextField("BirthdaysRange", {
                  label: "Days Range(numbers): ",
                }),
                PropertyPaneTextField("BackgroundImage", {
                  label: "Background Image Link: ",
                }),
                PropertyFieldNumber("pageItemsNumber", {
                  key: "pageItemsNumber",
                  label: "pageItemsNumber",
                  description: "pageItemsNumber",
                  value: this.properties.pageItemsNumber,
                  disabled: false
                }),
                PropertyFieldListPicker("SpotlightId", {
                  label: "Select Spotlight list",
                  selectedList: this.properties.SpotlightId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                  key: "SpotlightId",
                }),
                PropertyFieldNumber("SpotlightInterval", {
                  key: "SpotlightInterval",
                  label: "Spotlight Interval",
                  description: "Spotlight Interval",
                  value: this.properties.SpotlightInterval,
                  disabled: false
                }),

              ]
            }
          ]
        }
      ]
    };
  }
}
