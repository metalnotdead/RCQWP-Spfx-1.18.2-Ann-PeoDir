import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'AnniversariesWebPartStrings';
import Anniversaries from './components/Anniversaries';
import { IAnniversariesProps } from './components/IAnniversariesProps';
import { getSP } from '../pnpjsConfig';

export interface IAnniversariesWebPartProps {
  description: string;
  title: string;
  pageSize: number;
  dateField: string;
  dateFieldAs: string;
  daysFromTodayFilter: number;
  daysBeforeTodayFilter: number;
  textField: string;
  secondaryTextField: string;
  tertiaryTextField: string;
  personaSize: number;
  noResultsMessage: string;
  celebrateIcon: string;
  filterField: string;
  additionalFilterKQL: string;
}

export default class AnniversariesWebPart extends BaseClientSideWebPart<IAnniversariesWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IAnniversariesProps> = React.createElement(
      Anniversaries,
      {
        displayMode: this.displayMode,
        webUrl: this.context.pageContext.web.absoluteUrl,
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        title: this.properties.title,
        pageSize: this.properties.pageSize,
        dateField: this.properties.dateField,
        dateFieldAs:this.properties.dateFieldAs,
        daysFromTodayFilter: this.properties.daysFromTodayFilter,
        daysBeforeTodayFilter: this.properties.daysBeforeTodayFilter,
        textField: this.properties.textField,
        secondaryTextField: this.properties.secondaryTextField,
        tertiaryTextField: this.properties.tertiaryTextField,
        personaSize: this.properties.personaSize,
        noResultsMessage: this.properties.noResultsMessage,
        celebrateIcon: this.properties.celebrateIcon,
        filterField: this.properties.filterField,
        additionalFilterKQL: this.properties.additionalFilterKQL,
        onTitleUpdate: (newTitle: string) => {
					// after updating the web part title in the component
					// persist it in web part properties
					this.properties.title = newTitle;
				}
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    getSP(this.context);
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
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
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
    return Version.parse('1.0');
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected onAfterPropertyPaneChangesApplied(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
    this.render();
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
              groupName: strings.SettingsLabel,
              groupFields: [
                PropertyPaneSlider('pageSize', {
                  label: strings.PageSizeFieldLabel,
                  min: 1,
                  max: 20,
                  value: 10
                }),
                PropertyPaneSlider('personaSize', {
                  label: strings.PersonaSizeFieldLabel,
                  min: 0,
                  max: 18,
                  value: 14
                }),
                PropertyPaneTextField('noResultsMessage', {
                  label: strings.NoResultsMessageLabel,
                }),
                PropertyPaneTextField('celebrateIcon', {
                  label: strings.CelebrateIconLabel,
                  description: strings.CelebrateIconDescription
                }),
              ]
            },
            {
              groupName: strings.FilterSettingsLabel,
              groupFields: [
                PropertyPaneTextField('filterField', {
                  label: strings.FilterFieldLabel
                }),
                PropertyPaneDropdown('dateFieldAs', {
                  label: strings.FilterAsLabel,
                  options: [
                    { key: '0', text: 'as is' },
                    { key: '1', text: 'as 2000`s' },
                  ], 
                  selectedKey: '1'
                }),
                PropertyPaneTextField('additionalFilterKQL', {
                  label: strings.AdditionalFilterFieldLabel,
                  description: strings.AdditionalFilterFieldDescription
                }),
                PropertyPaneSlider('daysFromTodayFilter', {
                  label: strings.DaysFromTodayFilterFieldLabel,
                  min: 1,
                  max: 365,
                  value: 7
                }),
                PropertyPaneSlider('daysBeforeTodayFilter', {
                  label: strings.DaysBeforeTodayFilterFieldLabel,
                  min: -365,
                  max: 0,
                  value: 0
                }),
              ]
            },
            {
              groupName: strings.FieldMappingLabel,
              groupFields: [
                PropertyPaneTextField('dateField', {
                  label: strings.DateFieldLabel
                }),
                PropertyPaneTextField('textField', {
                  label: strings.TextFieldLabel
                }),
                PropertyPaneTextField('secondaryTextField', {
                  label: strings.SecondaryTextFieldLabel
                }),
                PropertyPaneTextField('tertiaryTextField', {
                  label: strings.TertiaryTextFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
