import { DisplayMode } from '@microsoft/sp-core-library';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from '@microsoft/sp-component-base';

export interface ICalendarProps {
  title: string;
  feedsTitle: string;
  siteUrl: string;
  list: string;
  siteUrl2: string;
  list2: string;
  displayMode: DisplayMode;
  updateTitleProperty: (value: string) => void;
  updateFeedsTitleProperty: (value: string) => void;
  context: WebPartContext;
  headerColor:string;
  calendarCellColor: string;
  themeVariant: IReadonlyTheme;
  headingTitleColor: string;
}
