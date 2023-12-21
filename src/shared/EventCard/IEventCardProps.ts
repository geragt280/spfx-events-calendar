import { ICalendarEvent } from "../../shared/CalendarService";
import { IReadonlyTheme } from '@microsoft/sp-component-base';

export interface IEventCardProps {
    isEditMode: boolean;
    event: ICalendarEvent;
    isNarrow: boolean;
    themeVariant?: IReadonlyTheme;
}
