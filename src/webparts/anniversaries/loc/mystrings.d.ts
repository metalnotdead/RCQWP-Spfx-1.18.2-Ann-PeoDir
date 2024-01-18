declare interface IAnniversariesWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  //Custom propeties start here
  SettingsLabel: string;
  FieldMappingLabel: string;
  TitleFieldLabel : string;
  PageSizeFieldLabel: string;
  DateFieldLabel: string;
  DaysFromTodayFilterFieldLabel : string;
  DaysBeforeTodayFilterFieldLabel: string;
  TextFieldLabel: string;
  TertiaryTextFieldLabel: string;
  SecondaryTextFieldLabel: string;
  PersonaSizeFieldLabel:string;
  FilterSettingsLabel:string;
  NoResultsMessageLabel:string;
  CelebrateIconLabel:string;
  CelebrateIconDescription:string;
  FilterFieldLabel:string;
  FilterAsLabel:string;
  AdditionalFilterFieldLabel:string;
  AdditionalFilterFieldDescription:string;
}

declare module 'AnniversariesWebPartStrings' {
  const strings: IAnniversariesWebPartStrings;
  export = strings;
}
