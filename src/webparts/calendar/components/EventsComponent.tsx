import { DisplayMode } from "@microsoft/sp-core-library";
import { FocusZone, FocusZoneDirection, List } from "office-ui-fabric-react";
import {
  CalendarServiceProviderType,
  ICalendarEvent,
  ICalendarService,
} from "../../../shared/CalendarService";
import { EventCard } from "../../../shared/EventCard";
import { Pagination } from "../../../shared/Pagination";
import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

export interface ICalendarFeedSummaryState {
  events: ICalendarEvent[];
  error: any | undefined;
  isLoading: boolean;
  currentPage: number;
}

export interface ICalendarFeedSummaryProps {
  title: string;
  displayMode: DisplayMode;
  context: WebPartContext;
  updateProperty: (value: string) => void;
  isConfigured: boolean;
  provider: ICalendarService;
  maxEvents: number;
  themeVariant: IReadonlyTheme;
  clientWidth: number;
}

export default function EventsComponent() {

    
  const _renderNarrowList = (): JSX.Element => {
    const { events, currentPage } = this.state;

    const { maxEvents } = this.props;

    // if we're in edit mode, let's not make the events clickable
    const isEditMode: boolean = this.props.displayMode === DisplayMode.Edit;

    let pagedEvents: ICalendarEvent[] = events;
    let usePaging: boolean = false;

    if (maxEvents > 0 && events.length > maxEvents) {
      // calculate the page size
      const pageStartAt: number = maxEvents * (currentPage - 1);
      const pageEndAt: number = maxEvents * currentPage;

      pagedEvents = events.slice(pageStartAt, pageEndAt);
      usePaging = true;
    }

    return (
      <FocusZone
        direction={FocusZoneDirection.vertical}
        isCircularNavigation={false}
        data-automation-id={"narrow-list"}
        aria-label={isEditMode ? "Edit Mode Enable" : "Edit Mode Disable"}
      >
        <List
          items={pagedEvents}
          onRenderCell={(item, _index) => (
            <EventCard
              isEditMode={isEditMode}
              event={item}
              isNarrow={true}
              themeVariant={this.props.themeVariant}
            />
          )}
        />
        {usePaging && (
          <Pagination
            showPageNum={false}
            currentPage={currentPage}
            itemsCountPerPage={maxEvents}
            totalItems={events.length}
            onPageUpdate={this._onPageUpdate}
          />
        )}
      </FocusZone>
    );
  };

  return _renderNarrowList();
}
