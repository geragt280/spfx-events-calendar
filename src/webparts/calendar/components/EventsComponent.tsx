import { DisplayMode } from "@microsoft/sp-core-library";
import { css, FocusZone, FocusZoneDirection, List, Spinner } from "office-ui-fabric-react";
import {
  ICalendarEvent,
} from "../../../shared/CalendarService";
import { EventCard } from "../../../shared/EventCard";
import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
// import moment from "moment";
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import styles from "./CalendarFeedSummary.module.scss";

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
  feedEvents: ICalendarEvent[];
  isLoading: boolean;
  maxEvents: number;
  clientWidth: number;
}

export default class EventsComponent extends React.Component<ICalendarFeedSummaryProps, ICalendarFeedSummaryState> {
  private MaxMobileWidth: number = 480;
  // Your component logic goes here
  constructor(props: ICalendarFeedSummaryProps) {
    super(props);
    this.state = {
      isLoading: false,
      events: props.feedEvents,
      error: undefined,
      currentPage: 1
    };
  }

  public componentDidMount(): void {
    if (this.props.isConfigured) {
      this._loadEvents();
    }
  }

  public componentDidUpdate(prevProps: ICalendarFeedSummaryProps, prevState: ICalendarFeedSummaryState): void {
    const { feedEvents, isLoading } = this.props;
    // if we didn't have a provider and now we do, we definitely need to update
    if (!isLoading) {
      if (feedEvents.length > 0) {
        this._loadEvents();
      }

      // there's nothing to do because there isn't a provider
      return;
    }
  }

  private _loadEvents(): void {
    const { isLoading, feedEvents, maxEvents } = this.props;
    
    if (isLoading) {
      this.setState({
        isLoading: true
      });

      try {
        let events = feedEvents;
        if (maxEvents > 0) {
          events = events.slice(0, maxEvents);
        }
        // don't cache in the case of errors
        this.setState({
          isLoading: false,
          error: undefined,
          events: events
        });
        return;
      }
      catch (error) {
        console.log("Exception returned by getEvents", error.message);
        this.setState({
          isLoading: false,
          error: error.message,
          events: []
        });
      }
    }
  }

  private _onPageUpdate = (pageNumber: number): void => {
    this.setState({
      currentPage: pageNumber
    });
  }

  private _renderNarrowList = (): JSX.Element => {
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
              themeVariant={null}
            />
          )}
        />
      </FocusZone>
    );
  }

  private _onConfigure = () => {
    this.props.context.propertyPane.open();
  }

  private _renderContent(): JSX.Element {
    const {
      displayMode,
    } = this.props;
    const {
      events,
      isLoading,
      error
    } = this.state;

    const isEditMode: boolean = displayMode === DisplayMode.Edit;
    const hasErrors: boolean = error !== undefined;
    const hasEvents: boolean = events.length > 0;

    if (isLoading) {
      // we're currently loading
      return (
      <div className={styles.spinner}>
        <Spinner label={"Please wait..."} />
      </div>);
    }

    if (!hasEvents) {
      // we're done loading, no errors, but have no events
      return (<div className={styles.emptyMessage}>{"There aren't any upcoming events."}</div>);
    }

    return this._renderNarrowList();
  }

  public render(): React.ReactElement<ICalendarFeedSummaryProps> {
    const {
      isConfigured,
    } = this.props;

    // if we're not configured, show the placeholder
    if (!isConfigured) {
      return (
        <Placeholder
          iconName="Calendar"
          iconText={"Configure event feed"}
          description={"To display a summary of events, you need to select a feed type and configure the event feed URL."}
          buttonLabel={"Configure"}
          onConfigure={this._onConfigure} />
      );
    }

    // we're configured, let's show stuff

    // put everything together in a nice little calendar view
    return (
      <div className={css(styles.calendarFeedSummary, styles.webPartChrome)} style={{ backgroundColor: 'wheat' }}>
        <div className={css(styles.webPartHeader, styles.headerSmMargin)}>
          <WebPartTitle displayMode={this.props.displayMode}
            title={this.props.title}
            updateProperty={this.props.updateProperty}
          />
        </div>
        <div className={styles.content}>
          {this._renderContent()}
        </div>
      </div>
    );
  }
}

