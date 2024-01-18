import { DisplayMode } from "@microsoft/sp-core-library";

export interface IAnniversariesProps {
  webUrl: string;
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  title: string;
  pageSize: number;
  personaSize:number;
  dateField: string;
  dateFieldAs: string;
  daysFromTodayFilter: number;
  daysBeforeTodayFilter: number;
  textField: string;
  secondaryTextField: string;
  tertiaryTextField: string;
  noResultsMessage: string;
  celebrateIcon:string;
  filterField:string;
  additionalFilterKQL:string;
  /**
   * Current page display mode. Used to determine if the user should
   * be able to edit the page title or not.
   */
  displayMode: DisplayMode;
  /**
   * Event handler for changing the web part title
   */
  onTitleUpdate: (newTitle: string) => void;
}
