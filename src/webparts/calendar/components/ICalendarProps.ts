import { DisplayMode } from '@microsoft/sp-core-library';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IDateTimeFieldValue } from '@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

export interface ICalendarProps {
  title: string;
  feedsTitle: string;
  siteUrl: string;
  list: string;
  eventStartDate:  IDateTimeFieldValue;
  eventEndDate: IDateTimeFieldValue;
  siteUrl2: string;
  list2: string;
  eventStartDate2:  IDateTimeFieldValue;
  eventEndDate2: IDateTimeFieldValue;
  displayMode: DisplayMode;
  updateTitleProperty: (value: string) => void;
  updateFeedsTitleProperty: (value: string) => void;
  context: WebPartContext;
  headerColor:string;
  calendarCellColor: string;
  themeVariant: IReadonlyTheme;
  headingTitleColor: string;
}
