import { DisplayMode } from "@microsoft/sp-core-library";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPFI } from "@pnp/sp";

export interface ICalendarProps {
  title: string;
  feedsTitle: string;
  list: string;
  list1column1: string;
  list1column2: string;
  list1column3: string;
  list1column4: string;  
  list1column5: string;
  list2: string;
  list2column1: string;
  list2column2: string;
  list2column3: string;
  list2column4: string;
  list2column5: string;
  displayMode: DisplayMode;
  updateTitleProperty: (value: string) => void;
  updateFeedsTitleProperty: (value: string) => void;
  context: WebPartContext;
  headerColor:string;
  calendarCellColor: string;
  themeVariant: IReadonlyTheme;
  headingTitleColor: string;
  showEventsFeedsWP: boolean;
  sp: SPFI;
}
