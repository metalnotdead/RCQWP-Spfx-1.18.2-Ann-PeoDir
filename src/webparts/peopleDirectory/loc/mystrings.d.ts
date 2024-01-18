declare interface IPeopleDirectoryWebPartStrings {
  SearchButtonText: string;
  LoadingSpinnerLabel: string;
  NoPeopleFoundLabel: string;
  SearchBoxPlaceholder: string;
  ErrorLabel: string;
  SkillsLabel: string;
  ProjectsLabel: string;
  CopyEmailLabel: string;
  CopyPhoneLabel: string;
  CopyMobileLabel: string;
  AdditionalFilterFieldLabel: string;
  AdditionalFilterFieldDescription: string;
  PropertyPaneDescription:string;
  FilterSettingsLabel:string;
  SettingsLabel:string;
  TitleFieldLabel:string;
  IndexByLastname:string;
}

declare module 'PeopleDirectoryWebPartStrings' {
  const strings: IPeopleDirectoryWebPartStrings;
  export = strings;
}
