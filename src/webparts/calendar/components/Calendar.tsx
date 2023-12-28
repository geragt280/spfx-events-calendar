import * as React from "react";
import styles from "./Calendar.module.scss";
import { ICalendarProps } from "./ICalendarProps";
import { escape } from "@microsoft/sp-lodash-subset";
// import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import * as moment from "moment";
// import * as strings from "CalendarWebPartStrings";
import "react-big-calendar/lib/css/react-big-calendar.css";
// import FullCalendar from "@fullcalendar/react";
// import dayGridPlugin from "@fullcalendar/daygrid"; // a plugin!
// require("./calendar.css");
import {
  css,
  FontIcon,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Text,
} from "office-ui-fabric-react";
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { IDateTimeFieldValue } from "@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker";
import {
  Calendar as MyCalendar,
  DayPropGetter,
  EventPropGetter,
  // EventWrapperProps,
  momentLocalizer,
  ToolbarProps,
} from "react-big-calendar";
import { IEventData } from "./Interfaces/IEventData";
import { ICalendarEvent } from "./Interfaces/ICalendarEvent";
import { IUserPermissions } from "./Interfaces/IUserPermissions";
import { PermissionKind } from "@pnp/sp/security";
// import parseRecurrentEvent from "./services/parseRecurrentEvent";
import { DisplayMode } from "@microsoft/sp-core-library";
import EventsComponent from "./EventsComponent";
// import CustomEvent from "./CustomEvent";

interface ICalendarState {
  showDialog: boolean;
  eventData: IEventData[];
  selectedEvent: IEventData | null;
  startDateSlot?: Date;
  endDateSlot?: Date;
  isloading: boolean;
  hasError: boolean;
  errorMessage: string;
  feedsEvents: ICalendarEvent[];
  monthAddCount: number;
  calenderIsLoading: boolean;
}

const localizer = momentLocalizer(moment);

export default class Calendar extends React.Component<
  ICalendarProps,
  ICalendarState
> {
  private userListPermissions: IUserPermissions = undefined;
  private siteUrl: string = this.props.context.pageContext.web.absoluteUrl;
  private eventStartDate: IDateTimeFieldValue = {
    value: moment().startOf("month").subtract(7, "days").toDate(),
    displayValue: moment().format("ddd MMM MM YYYY"),
  };
  private eventEndDate: IDateTimeFieldValue = {
    value: moment().endOf("month").add(7, "days").toDate(),
    displayValue: moment().format("ddd MMM MM YYYY"),
  };
  private eventStartDate2: IDateTimeFieldValue = {
    value: moment().add(1, "days").toDate(),
    displayValue: moment().format("ddd MMM MM YYYY"),
  };
  private eventEndDate2: IDateTimeFieldValue = {
    value: moment().add(5, "years").endOf("month").toDate(),
    displayValue: moment().format("ddd MMM MM YYYY"),
  };

  public constructor(props: ICalendarProps) {
    super(props);

    this.state = {
      showDialog: false,
      eventData: [],
      selectedEvent: undefined,
      isloading: true,
      hasError: false,
      errorMessage: "",
      feedsEvents: [],
      monthAddCount: 0,
      calenderIsLoading: false,
    };

    this.onDismissPanel = this.onDismissPanel.bind(this);
    this.onSelectEvent = this.onSelectEvent.bind(this);
    this.onSelectSlot = this.onSelectSlot.bind(this);
    moment.locale(
      this.props.context.pageContext.cultureInfo.currentUICultureName
    );
  }

  public async componentDidMount(): Promise<void> {
    if (DisplayMode.Edit === this.props.displayMode) {
      console.log("context urls", this.props.context.pageContext.web);
    }

    await this.loadWebpart();
  }

  private loadWebpart = async (): Promise<void> => {
    this.setState({ isloading: true });
    await this.loadEvents();
    this.setState({ isloading: false });
  };

  private async onDismissPanel(refresh: boolean) {
    this.setState({ showDialog: false });
    if (refresh === true) {
      this.setState({ isloading: true });
      await this.loadEvents();
      this.setState({ isloading: false });
    }
  }

  private onSelectEvent(event: any) {
    this.setState({
      showDialog: true,
      selectedEvent: event,
    });
  }

  public async onSelectSlot({ start, end }: { start: any; end: any }) {
    if (!this.userListPermissions.hasPermissionAdd) return;
    this.setState({
      showDialog: true,
      startDateSlot: start,
      endDateSlot: end,
      selectedEvent: undefined,
    });
  }

  public async getLocalTime(date: string | Date): Promise<string> {
    const { sp } = this.props;
    try {
      const localTime = await sp.web.regionalSettings.timeZone.utcToLocalTime(
        date
      );
      return localTime;
    } catch (error) {
      return Promise.reject(error);
    }
  }

  public async getUserPermissions(
    siteUrl: string,
    listId: string
  ): Promise<IUserPermissions> {
    const { sp } = this.props;
    let hasPermissionAdd: boolean = false;
    let hasPermissionEdit: boolean = false;
    let hasPermissionDelete: boolean = false;
    let hasPermissionView: boolean = false;
    let userPermissions: IUserPermissions = undefined;
    try {
      const web = sp.web;
      const userEffectivePermissions = await web.lists
        .getById(listId)
        .effectiveBasePermissions();
      // ...
      hasPermissionAdd = sp.web.lists
        .getById(listId)
        .hasPermissions(userEffectivePermissions, PermissionKind.AddListItems);
      hasPermissionDelete = sp.web.lists
        .getById(listId)
        .hasPermissions(
          userEffectivePermissions,
          PermissionKind.DeleteListItems
        );
      hasPermissionEdit = sp.web.lists
        .getById(listId)
        .hasPermissions(userEffectivePermissions, PermissionKind.EditListItems);
      hasPermissionView = sp.web.lists
        .getById(listId)
        .hasPermissions(userEffectivePermissions, PermissionKind.ViewListItems);
      userPermissions = {
        hasPermissionAdd: hasPermissionAdd,
        hasPermissionEdit: hasPermissionEdit,
        hasPermissionDelete: hasPermissionDelete,
        hasPermissionView: hasPermissionView,
      };
    } catch (error) {
      return Promise.reject(error);
    }
    return userPermissions;
  }

  public async getChoiceFieldOptions(
    siteUrl: string,
    listId: string,
    fieldInternalName: string
  ): Promise<{ key: string; text: string }[]> {
    let fieldOptions: { key: string; text: string }[] = [];
    const { sp } = this.props;
    try {
      const web = sp.web;
      const results = await web.lists
        .getById(listId)
        .fields.getByInternalNameOrTitle(fieldInternalName)
        .select("Title", "InternalName", "Choices")();
      if (results && results.Choices.length > 0) {
        for (const option of results.Choices) {
          fieldOptions.push({
            key: option,
            text: option,
          });
        }
      }
    } catch (error) {
      return Promise.reject(error);
    }
    return fieldOptions;
  }

  public async colorGenerate() {
    var hexValues = [
      "0",
      "1",
      "2",
      "3",
      "4",
      "5",
      "6",
      "7",
      "8",
      "9",
      "a",
      "b",
      "c",
      "d",
      "e",
    ];
    var newColor = "#";

    for (var i = 0; i < 6; i++) {
      var x = Math.round(Math.random() * 14);

      var y = hexValues[x];
      newColor += y;
    }
    return newColor;
  }

  public async getUserProfilePictureUrl(loginName: string) {
    let results: any = null;
    const { sp } = this.props;
    try {
      results = await sp.profiles.getPropertiesFor(loginName);
    } catch (error) {
      results = null;
    }
    return results.PictureUrl;
  }

  public async deCodeHtmlEntities(string: string) {
    const HtmlEntitiesMap = {
      "'": "&#39;",
      "<": "&lt;",
      ">": "&gt;",
      " ": "&nbsp;",
      "¡": "&iexcl;",
      "¢": "&cent;",
      "£": "&pound;",
      "¤": "&curren;",
      "¥": "&yen;",
      "¦": "&brvbar;",
      "§": "&sect;",
      "¨": "&uml;",
      "©": "&copy;",
      ª: "&ordf;",
      "«": "&laquo;",
      "¬": "&not;",
      "®": "&reg;",
      "¯": "&macr;",
      "°": "&deg;",
      "±": "&plusmn;",
      "²": "&sup2;",
      "³": "&sup3;",
      "´": "&acute;",
      µ: "&micro;",
      "¶": "&para;",
      "·": "&middot;",
      "¸": "&cedil;",
      "¹": "&sup1;",
      º: "&ordm;",
      "»": "&raquo;",
      "¼": "&frac14;",
      "½": "&frac12;",
      "¾": "&frac34;",
      "¿": "&iquest;",
      À: "&Agrave;",
      Á: "&Aacute;",
      Â: "&Acirc;",
      Ã: "&Atilde;",
      Ä: "&Auml;",
      Å: "&Aring;",
      Æ: "&AElig;",
      Ç: "&Ccedil;",
      È: "&Egrave;",
      É: "&Eacute;",
      Ê: "&Ecirc;",
      Ë: "&Euml;",
      Ì: "&Igrave;",
      Í: "&Iacute;",
      Î: "&Icirc;",
      Ï: "&Iuml;",
      Ð: "&ETH;",
      Ñ: "&Ntilde;",
      Ò: "&Ograve;",
      Ó: "&Oacute;",
      Ô: "&Ocirc;",
      Õ: "&Otilde;",
      Ö: "&Ouml;",
      "×": "&times;",
      Ø: "&Oslash;",
      Ù: "&Ugrave;",
      Ú: "&Uacute;",
      Û: "&Ucirc;",
      Ü: "&Uuml;",
      Ý: "&Yacute;",
      Þ: "&THORN;",
      ß: "&szlig;",
      à: "&agrave;",
      á: "&aacute;",
      â: "&acirc;",
      ã: "&atilde;",
      ä: "&auml;",
      å: "&aring;",
      æ: "&aelig;",
      ç: "&ccedil;",
      è: "&egrave;",
      é: "&eacute;",
      ê: "&ecirc;",
      ë: "&euml;",
      ì: "&igrave;",
      í: "&iacute;",
      î: "&icirc;",
      ï: "&iuml;",
      ð: "&eth;",
      ñ: "&ntilde;",
      ò: "&ograve;",
      ó: "&oacute;",
      ô: "&ocirc;",
      õ: "&otilde;",
      ö: "&ouml;",
      "÷": "&divide;",
      ø: "&oslash;",
      ù: "&ugrave;",
      ú: "&uacute;",
      û: "&ucirc;",
      ü: "&uuml;",
      ý: "&yacute;",
      þ: "&thorn;",
      ÿ: "&yuml;",
      Œ: "&OElig;",
      œ: "&oelig;",
      Š: "&Scaron;",
      š: "&scaron;",
      Ÿ: "&Yuml;",
      ƒ: "&fnof;",
      ˆ: "&circ;",
      "˜": "&tilde;",
      Α: "&Alpha;",
      Β: "&Beta;",
      Γ: "&Gamma;",
      Δ: "&Delta;",
      Ε: "&Epsilon;",
      Ζ: "&Zeta;",
      Η: "&Eta;",
      Θ: "&Theta;",
      Ι: "&Iota;",
      Κ: "&Kappa;",
      Λ: "&Lambda;",
      Μ: "&Mu;",
      Ν: "&Nu;",
      Ξ: "&Xi;",
      Ο: "&Omicron;",
      Π: "&Pi;",
      Ρ: "&Rho;",
      Σ: "&Sigma;",
      Τ: "&Tau;",
      Υ: "&Upsilon;",
      Φ: "&Phi;",
      Χ: "&Chi;",
      Ψ: "&Psi;",
      Ω: "&Omega;",
      α: "&alpha;",
      β: "&beta;",
      γ: "&gamma;",
      δ: "&delta;",
      ε: "&epsilon;",
      ζ: "&zeta;",
      η: "&eta;",
      θ: "&theta;",
      ι: "&iota;",
      κ: "&kappa;",
      λ: "&lambda;",
      μ: "&mu;",
      ν: "&nu;",
      ξ: "&xi;",
      ο: "&omicron;",
      π: "&pi;",
      ρ: "&rho;",
      ς: "&sigmaf;",
      σ: "&sigma;",
      τ: "&tau;",
      υ: "&upsilon;",
      φ: "&phi;",
      χ: "&chi;",
      ψ: "&psi;",
      ω: "&omega;",
      ϑ: "&thetasym;",
      ϒ: "&Upsih;",
      ϖ: "&piv;",
      "–": "&ndash;",
      "—": "&mdash;",
      "‘": "&lsquo;",
      "’": "&rsquo;",
      "‚": "&sbquo;",
      "“": "&ldquo;",
      "”": "&rdquo;",
      "„": "&bdquo;",
      "†": "&dagger;",
      "‡": "&Dagger;",
      "•": "&bull;",
      "…": "&hellip;",
      "‰": "&permil;",
      "′": "&prime;",
      "″": "&Prime;",
      "‹": "&lsaquo;",
      "›": "&rsaquo;",
      "‾": "&oline;",
      "⁄": "&frasl;",
      "€": "&euro;",
      ℑ: "&image;",
      "℘": "&weierp;",
      ℜ: "&real;",
      "™": "&trade;",
      ℵ: "&alefsym;",
      "←": "&larr;",
      "↑": "&uarr;",
      "→": "&rarr;",
      "↓": "&darr;",
      "↔": "&harr;",
      "↵": "&crarr;",
      "⇐": "&lArr;",
      "⇑": "&UArr;",
      "⇒": "&rArr;",
      "⇓": "&dArr;",
      "⇔": "&hArr;",
      "∀": "&forall;",
      "∂": "&part;",
      "∃": "&exist;",
      "∅": "&empty;",
      "∇": "&nabla;",
      "∈": "&isin;",
      "∉": "&notin;",
      "∋": "&ni;",
      "∏": "&prod;",
      "∑": "&sum;",
      "−": "&minus;",
      "∗": "&lowast;",
      "√": "&radic;",
      "∝": "&prop;",
      "∞": "&infin;",
      "∠": "&ang;",
      "∧": "&and;",
      "∨": "&or;",
      "∩": "&cap;",
      "∪": "&cup;",
      "∫": "&int;",
      "∴": "&there4;",
      "∼": "&sim;",
      "≅": "&cong;",
      "≈": "&asymp;",
      "≠": "&ne;",
      "≡": "&equiv;",
      "≤": "&le;",
      "≥": "&ge;",
      "⊂": "&sub;",
      "⊃": "&sup;",
      "⊄": "&nsub;",
      "⊆": "&sube;",
      "⊇": "&supe;",
      "⊕": "&oplus;",
      "⊗": "&otimes;",
      "⊥": "&perp;",
      "⋅": "&sdot;",
      "⌈": "&lceil;",
      "⌉": "&rceil;",
      "⌊": "&lfloor;",
      "⌋": "&rfloor;",
      "⟨": "&lang;",
      "⟩": "&rang;",
      "◊": "&loz;",
      "♠": "&spades;",
      "♣": "&clubs;",
      "♥": "&hearts;",
      "♦": "&diams;",
    };

    var entityMap: any = HtmlEntitiesMap;
    for (var key in entityMap) {
      var entity = entityMap[key];
      var regex = new RegExp(entity, "g");
      string = string.replace(regex, key);
    }
    string = string.replace(/&quot;/g, '"');
    string = string.replace(/&amp;/g, "&");
    return string;
  }

  public async getSpecificEvents(
    listId: string,
    eventStartDate: Date,
    eventEndDate: Date,
    titleColumnName: string,
    startDateColumnName: string,
    endDateColumnName: string,
    colorColumnName: string,
    allDayColumnName: string
  ): Promise<IEventData[]> {
    const { sp } = this.props;
    const { siteUrl } = this;
    let events: IEventData[] = [];
    if (!siteUrl) {
      return [];
    }
    try {
      const results = await sp.web.lists.getById(listId).items.filter(`${startDateColumnName} ge '${moment(eventStartDate).format('YYYY-MM-DD')}' and ${endDateColumnName} le '${moment(eventEndDate).format('YYYY-MM-DD')}'`).orderBy(startDateColumnName, true)();

      if (DisplayMode.Edit === this.props.displayMode) {
        console.log("Results", results);
      }

      await Promise.all(
        results.map(async (e) => {
          const isAllDay: boolean = e[allDayColumnName];
          const localStartDate = await this.getLocalTime(
            e[startDateColumnName]
          );
          const localEndDate = await this.getLocalTime(e[endDateColumnName]);
          // if (e[colorColumnName]) {
          //   debugger;
          // }
          events.push({
            title: e[titleColumnName],
            EventDate: new Date(localStartDate),
            EndDate: new Date(localEndDate),
            color: e[colorColumnName],
            fAllDayEvent: isAllDay,
          });
        })
      );

      // Return Data
      return events;
    } catch (error) {
      if (DisplayMode.Edit === this.props.displayMode) {
        console.dir(error);        
      }
      return Promise.reject(error);
    }
  }

  public async getAllEvents(
    listId: string,
    titleColumnName: string,
    startDateColumnName: string,
    endDateColumnName: string,
    colorColumnName: string,
    allDayColumnName: string
  ): Promise<IEventData[]> {
    const { sp } = this.props;
    const { siteUrl } = this;
    let events: IEventData[] = [];
    if (!siteUrl) {
      return [];
    }
    try {
      const results = await sp.web.lists.getById(listId).items.orderBy(startDateColumnName, true)();

      if (DisplayMode.Edit === this.props.displayMode) {
        console.log("Results", results);
      }

      await Promise.all(
        results.map(async (e) => {
          const isAllDay: boolean = e[allDayColumnName];
          const localStartDate = await this.getLocalTime(
            e[startDateColumnName]
          );
          const localEndDate = await this.getLocalTime(e[endDateColumnName]);
          // if (e[colorColumnName]) {
          //   debugger;
          // }
          events.push({
            title: e[titleColumnName],
            EventDate: new Date(localStartDate),
            EndDate: new Date(localEndDate),
            color: e[colorColumnName],
            fAllDayEvent: isAllDay,
            Description: e["Description"],
            Category: e["Category"]
          });
        })
      );

      // Return Data
      return events;
    } catch (error) {
      if (DisplayMode.Edit === this.props.displayMode) {
        console.dir(error);        
      }
      return Promise.reject(error);
    }
  }

  // public async getAllEvents(listId: string, eventStartDate: Date, eventEndDate: Date): Promise<IEventData[]> {
  //   const { siteUrl } = this;
  //   const { sp } = this.props;
  //   let events: IEventData[] = [];
  //   if (!siteUrl) {
  //     return [];
  //   }
  //   try {
  //     // Get Category Field Choices
  //     const categoryDropdownOption = await this.getChoiceFieldOptions(siteUrl, listId, 'Category');
  //     let categoryColor: { category: string, color: string }[] = [];
  //     for (const cat of categoryDropdownOption) {
  //       categoryColor.push({ category: cat.text, color: await this.colorGenerate() });
  //     }

  //     const web = sp.web;
  //     const results = await web.lists.getById(listId).renderListDataAsStream(
  //       {
  //         DatesInUtc: true,
  //         ViewXml: `<View><ViewFields><FieldRef Name='RecurrenceData'/><FieldRef Name='Duration'/><FieldRef Name='Author'/><FieldRef Name='Category'/><FieldRef Name='Description'/><FieldRef Name='ParticipantsPicker'/><FieldRef Name='Geolocation'/><FieldRef Name='ID'/><FieldRef Name='EndDate'/><FieldRef Name='EventDate'/><FieldRef Name='ID'/><FieldRef Name='Location'/><FieldRef Name='Title'/><FieldRef Name='fAllDayEvent'/><FieldRef Name='EventType'/><FieldRef Name='UID' /><FieldRef Name='fRecurrence' /></ViewFields>
  //         <Query>
  //         <Where>
  //           <And>
  //             <Geq>
  //               <FieldRef Name='EventDate' />
  //               <Value IncludeTimeValue='false' Type='DateTime'>${moment(eventStartDate).format('YYYY-MM-DD')}</Value>
  //             </Geq>
  //             <Leq>
  //               <FieldRef Name='EventDate' />
  //               <Value IncludeTimeValue='false' Type='DateTime'>${moment(eventEndDate).format('YYYY-MM-DD')}</Value>
  //             </Leq>
  //             </And>
  //         </Where>
  //         <OrderBy>
  //           <FieldRef Name='EventDate' Ascending='True' />
  //         </OrderBy>
  //         </Query>
  //         <RowLimit Paged=\"FALSE\">2000</RowLimit>
  //         </View>`
  //       }
  //     );

  //     if (results && results.Row.length > 0) {
  //       let event: any = '';
  //       const mapEvents = async () : Promise<boolean> => {
  //           for (event of results.Row) {
  //             const eventDate = await this.getLocalTime(event.EventDate);
  //             const endDate = await this.getLocalTime(event.EndDate);
  //             const initialsArray: string[] = event.Author[0].title.split(' ');
  //             const initials: string = initialsArray[0].charAt(0) + initialsArray[initialsArray.length - 1].charAt(0);
  //             const userPictureUrl = await this.getUserProfilePictureUrl(`i:0#.f|membership|${event.Author[0].email}`);
  //             const attendees: number[] = [];
  //             const first: number = event.Geolocation.indexOf('(') + 1;
  //             const last: number = event.Geolocation.indexOf(')');
  //             const geo = event.Geolocation.substring(first, last);
  //             const geolocation = geo.split(' ');
  //             const CategoryColorValue: any[] = categoryColor.filter((value) => {
  //               return value.category == event.Category;
  //             });
  //             const isAllDayEvent: boolean = event["fAllDayEvent.value"] === "1";

  //             for (const attendee of event.ParticipantsPicker) {
  //               attendees.push(parseInt(attendee.id));
  //             }

  //             events.push({
  //               Id: event.ID,
  //               ID: event.ID,
  //               EventType: event.EventType,
  //               title: await this.deCodeHtmlEntities(event.Title),
  //               Description: event.Description,
  //               EventDate: isAllDayEvent ? new Date(event.EventDate.slice(0, -1)) : new Date(eventDate),
  //               EndDate: isAllDayEvent ? new Date(event.EndDate.slice(0, -1)) : new Date(endDate),
  //               location: event.Location,
  //               ownerEmail: event.Author[0].email,
  //               ownerPhoto: userPictureUrl ?
  //                 `https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=${event.Author[0].email}&UA=0&size=HR96x96` : '',
  //               ownerInitial: initials,
  //               color: CategoryColorValue.length > 0 ? CategoryColorValue[0].color : '#1a75ff', // blue default
  //               ownerName: event.Author[0].title,
  //               attendes: attendees,
  //               fAllDayEvent: isAllDayEvent,
  //               geolocation: { Longitude: parseFloat(geolocation[0]), Latitude: parseFloat(geolocation[1]) },
  //               Category: event.Category,
  //               Duration: event.Duration,
  //               RecurrenceData: event.RecurrenceData ? await this.deCodeHtmlEntities(event.RecurrenceData) : "",
  //               fRecurrence: event.fRecurrence,
  //               RecurrenceID: event.RecurrenceID ? event.RecurrenceID : undefined,
  //               MasterSeriesItemID: event.MasterSeriesItemID,
  //               UID: event.UID.replace("{", "").replace("}", ""),
  //             });
  //           }
  //         return true;
  //       };
  //       //Checks to see if there are any results saved in local storage
  //       if(window.localStorage.getItem("eventResult")){
  //         //if there is a local version - compares it to the current version 
  //         if(window.localStorage.getItem("eventResult") === JSON.stringify(results)){
  //           //No update needed use current savedEvents
  //           events = JSON.parse(window.localStorage.getItem("calendarEventsWithLocalTime"));
  //         }else{
  //           //update local storage
  //           window.localStorage.setItem("eventResult", JSON.stringify(results));
  //           //when they are not equal then we loop through the results and maps them to IEventData
  //           /* tslint:disable:no-unused-expression */
  //           await mapEvents() ? window.localStorage.setItem("calendarEventsWithLocalTime", JSON.stringify(events)) : null;           
  //         }
  //       }else{
  //         //if there is no local storage of the events we create them
  //         window.localStorage.setItem("eventResult", JSON.stringify(results));
  //         //we also needs to map through the events the first time and save the mapped version to local storage
  //         await mapEvents() ? window.localStorage.setItem("calendarEventsWithLocalTime", JSON.stringify(events)) : null;           
  //       }
  //     }
  //     let parseEvt: parseRecurrentEvent = new parseRecurrentEvent();
  //     events = parseEvt.parseEvents(events, null, null);
       
  //     // Return Data
  //     return events;
  //   } catch (error) {
  //     console.dir(error);
  //     return Promise.reject(error);
  //   }
  // }

  public convertEventsDataToFeedsEventsData = ({
    calEvents,
  }: {
    calEvents: IEventData[];
  }): ICalendarEvent[] => {
    try {
      // Once we get the list, convert to calendar events
      let events: ICalendarEvent[] = calEvents.map((item: any) => {
        // let eventUrl: string = undefined; //combine(webUrl, "DispForm.aspx?ID=" + item.Id);
        // if (item.color) {
        //   debugger;
        // }
        const eventItem: ICalendarEvent = {
          title: item.title,
          start: item.EventDate,
          end: item.EndDate,
          backgroundColor: item.color === undefined || item.color === null ? "#000" : item.color,
          allDay: item.fAllDayEvent,
          category: undefined,
          description: undefined,
          location: undefined
        };
        return eventItem;
      });
      // Return the calendar items
      return events;
    } catch (error) {      
      if (DisplayMode.Edit === this.props.displayMode) {
        console.log("Exception caught by catch in SharePoint provider", error);
      }
      throw error;
    }
  };

  private onConfigure() {
    // Context of the web part
    this.props.context.propertyPane.open();
  }

  private async loadEvents(): Promise<void> {
    const {
      list,
      list2,
      displayMode,
      showEventsFeedsWP,
      list1column1,
      list1column2,
      list1column3,
      list1column4,
      list1column5,
      list2column1,
      list2column2,
      list2column3,
      list2column4,
      list2column5,
    } = this.props;
    const { siteUrl } = this;

    try {
      // Teste Properties
      if (!list || !siteUrl) return;

      this.userListPermissions = await this.getUserPermissions(siteUrl, list);

      const eventsData: IEventData[] = await this.getAllEvents(
        escape(list),
        // this.eventStartDate.value,
        // this.eventEndDate.value
        list1column1,
        list1column2,
        list1column3,
        list1column4,
        list1column5
      );

      // const eventsDataNewFormat: ICalendarEvent[] = this.convertEventsDataToFeedsEventsData({
      //   calEvents: eventsData,
      // });

      if (DisplayMode.Edit === displayMode) {
        console.log("Events data", eventsData);
      }

      if (showEventsFeedsWP) {
        const feedEventsData: IEventData[] = await this.getSpecificEvents(
          escape(list2),
          this.eventStartDate2.value,
          this.eventEndDate2.value,
          list2column1,
          list2column2,
          list2column3,
          list2column4,
          list2column5
        );

        if (DisplayMode.Edit === displayMode) {
          console.log(
            "Feed Events data",
            feedEventsData,
            this.eventStartDate2.value
          );
        }

        const calendarFeedsEvents: ICalendarEvent[] = this.convertEventsDataToFeedsEventsData({
          calEvents: feedEventsData,
        });

        this.setState({
          eventData: eventsData,
          hasError: false,
          errorMessage: "",
          feedsEvents: calendarFeedsEvents,
        });
      } else {
        this.setState({
          eventData: eventsData,
          hasError: false,
          errorMessage: "",
        });
      }
    } catch (error) {
      console.error("Error in getItems", error);
      this.setState({
        hasError: true,
        errorMessage: error.message,
        isloading: false,
      });
    }
  }

  private MyCustomHeader: React.FC<ToolbarProps> = ({ label, onNavigate }) => {
    const { headerColor, list, displayMode,
      list1column1,
      list1column2,
      list1column3,
      list1column4,
      list1column5 } = this.props;

    const handlePrev = async () => {
      onNavigate('PREV');
      // this.setState({
      //   calenderIsLoading: true
      // });
      const currentAddCount = this.state.monthAddCount - 1;
      this.eventStartDate = { value: moment().add(currentAddCount, 'months').startOf('month').subtract(7,'days').toDate(), displayValue: moment().format('ddd MMM MM YYYY')};
      this.eventEndDate = { value: moment().add(currentAddCount, 'months').endOf('month').add(7,'days').toDate(), displayValue: moment().format('ddd MMM MM YYYY')};

      const eventsData: IEventData[] = await this.getAllEvents(
        // escape(siteUrl),
        escape(list),
        // this.eventStartDate.value,
        // this.eventEndDate.value
        list1column1,
        list1column2,
        list1column3,
        list1column4,
        list1column5
      );

      if (DisplayMode.Edit === displayMode) {
        console.log("Events data", this.eventStartDate, currentAddCount, this.eventEndDate, eventsData);
      }

      this.setState({
        eventData: eventsData,
        monthAddCount: currentAddCount,
        calenderIsLoading: false
      });

    };

    const handleNext = async () => {
      onNavigate('NEXT');
      // this.setState({
      //   calenderIsLoading: true
      // });
      const currentAddCount = this.state.monthAddCount + 1;
      this.eventStartDate = { value: moment().add(currentAddCount, 'months').startOf('month').subtract(7,'days').toDate(), displayValue: moment().format('ddd MMM MM YYYY')};
      this.eventEndDate = { value: moment().add(currentAddCount, 'months').endOf('month').add(7,'days').toDate(), displayValue: moment().format('ddd MMM MM YYYY')};

      const eventsData: IEventData[] = await this.getAllEvents(
        // escape(siteUrl),
        escape(list),
        // this.eventStartDate.value,
        // this.eventEndDate.value
        list1column1,
        list1column2,
        list1column3,
        list1column4,
        list1column5
      );

      if (DisplayMode.Edit === displayMode) {
        console.log("Events data", this.eventStartDate, this.eventEndDate, currentAddCount, eventsData);
      }


      this.setState({
        eventData: eventsData,
        monthAddCount: currentAddCount,
        calenderIsLoading: false
      });
    };

    return (
      <div style={{ backgroundColor: headerColor, textAlign: 'center', display:'flex', flexDirection:'row', justifyContent:'space-evenly', alignItems:'center', color:'#ffffff' }}>
        {/* <button onClick={handlePrev}>&lt; Prev</button> */}
        {!this.state.calenderIsLoading && <FontIcon aria-label="Compass" iconName="PageLeft" className={styles.iconStyle} onClick={handlePrev} />}
        <Text style={{color:'#fff', fontSize:18, padding: 10}}>{label}</Text>
        {!this.state.calenderIsLoading && <FontIcon aria-label="Compass" iconName="PageRight" className={styles.iconStyle} onClick={handleNext} />}
        {/* <button onClick={handleNext}>Next &gt;</button> */}
      </div>
    );
  }

  // private MyEventWrapper: React.FC<EventWrapperProps> = ({ children, event }) => {
  //   return (
  //     <div style={{ position: 'relative' }}>
  //       {children}
  //       {event && (
  //         <div
  //           style={{
  //             position: 'absolute',
  //             top: '50%',
  //             left: '50%',
  //             transform: 'translate(-50%, -50%)',
  //             width: '8px',
  //             height: '8px',
  //             borderRadius: '50%',
  //             backgroundColor: this.props.calendarCellColor // Change the color as needed
  //           }}
  //         />
  //       )}
  //     </div>
  //   );
  // };

  private dayPropGetter: DayPropGetter = (date: Date): { style?: React.CSSProperties; className?: string } => {
    const today = moment();
    const isToday = today.isSame(date, 'day');
  
    return {
      style: {
        backgroundColor: isToday ? this.props.headerColor : 'inherit',
      },
    };
  };

  public eventStyleGetter: EventPropGetter<IEventData> = (event, start, end, isSelected): any => {
    
    return {
      style: {
        backgroundColor: event.color,
        borderColor: "gray",
        color: "transparent",
        borderRadius: "10",
        display: "flex",
        width:8,
        alignItems: "center",
        justifyContent: "center",
        height: 10
      }
    };
  }

  private handleEventClick = (event: IEventData, e: React.SyntheticEvent) => {
    // Prevent the default behavior to avoid interference with the tooltip
    const {selectedEvent} = this.state;
    e.preventDefault();
    if (DisplayMode.Edit === this.props.displayMode) {
      console.log("Opened item", event);
    }
    this.setState({
      selectedEvent: selectedEvent === event ? null : event
    });
  };

  private handleCloseTooltip = () => {
    this.setState({
      selectedEvent: null
    });
  };

  public render(): React.ReactElement<ICalendarProps> {
    const { showEventsFeedsWP } = this.props;

    // console.error("Error");
    return (
      <div className={styles.calendar} style={{ backgroundColor: "white" }}>
        <div style={{ backgroundColor: this.props.headingTitleColor }}>
          <Text className={styles.webPartTitle}>{this.props.title}</Text>
          {/* <WebPartTitle
            displayMode={this.props.displayMode}
            title={this.props.title}
            className={styles.webPartTitle}
            updateProperty={this.props.updateTitleProperty}
          /> */}
        </div>
        {!this.props.list ||
        (!this.props.list2 && showEventsFeedsWP) ||
        !this.props.list1column1 ||
        !this.props.list1column2 ||
        !this.props.list1column3 ||
        !this.props.list1column4 ||
        !this.props.list1column5 ||
        !this.eventStartDate.value ||
        !this.eventEndDate.value ? (
          <Placeholder
            iconName="Edit"
            iconText={"Configure your Calendar Web Part"}
            description={"Please configure list calendar "}
            buttonLabel={"Configure"}
            hideButton={this.props.displayMode === DisplayMode.Read}
            onConfigure={this.onConfigure.bind(this)}
          />
        ) : // test if has errors
        this.state.hasError ? (
          <MessageBar messageBarType={MessageBarType.error}>
            {this.state.errorMessage}
          </MessageBar>
        ) : (
          // show Calendar
          // Test if is loading Events
          <div>
            {this.state.isloading ? (
              <Spinner size={SpinnerSize.large} label={"Loading events..."} />
            ) : (
              <div className={css(styles.container, styles.msGrid)}>
                <div className={styles.msGridRow}>
                  <div className={showEventsFeedsWP ? styles.msGridCol : ""}>
                    <MyCalendar
                        dayPropGetter={this.dayPropGetter.bind(this)}
                        localizer={localizer}
                        selectable
                        events={this.state.eventData}
                        startAccessor="EventDate"
                        endAccessor="EndDate"
                        eventPropGetter={this.eventStyleGetter.bind(this)}
                        onSelectSlot={this.onSelectSlot.bind(this)}
                        defaultView="month"
                        view="month"
                        views={["month"]}
                        allDayMaxRows={1}
                        popup={true}
                        style={{ minHeight: 350 }}
                        onSelectEvent={this.handleEventClick.bind(this)}
                        components={{
                          toolbar: this.MyCustomHeader.bind(this),
                          // eventWrapper: this.MyEventWrapper.bind(this),
                          // event: CustomEvent
                        }}
                        defaultDate={moment().startOf("day").toDate()}
                      />
                    {/* <FullCalendar
                      plugins={[dayGridPlugin]}
                      initialView="dayGridMonth"
                      events={this.state.eventData}
                      height='auto'
                      headerToolbar={{
                        center: "title",
                        start: "",
                        left: "prev",
                        end: "",
                        right: "next",
                      }}
                      dayMaxEventRows={1}                      
                    /> */}
                  </div>
                  {showEventsFeedsWP && (
                    <div className={styles.msGridCol}>
                      <EventsComponent
                        context={this.props.context}
                        displayMode={this.props.displayMode}
                        clientWidth={400}
                        isConfigured={true}
                        maxEvents={4}
                        isLoading={this.state.isloading}
                        title={this.props.feedsTitle}
                        feedEvents={this.state.feedsEvents}
                        updateProperty={this.props.updateFeedsTitleProperty}
                      />
                    </div>
                  )}
                </div>
              </div>
            )}
          </div>
        )}
        {this.state.selectedEvent && (
          <div className="tooltipa" style={
            { 
              position:'absolute',
              top: 195,
              left: 200,
              backgroundColor: this.props.headerColor,
              height: 100,
              zIndex: 1000,
              padding:10,
              maxWidth:300,
              minWidth:150
            }
          }>
            <div>
              <FontIcon iconName="Cancel" className={styles.iconStyle} onClick={this.handleCloseTooltip} style={{float:'right'}} />
            </div>
            <div>
              <Text style={{fontWeight:'bold'}}>{this.state.selectedEvent.title}</Text>
            </div>
            <div>
              <p dangerouslySetInnerHTML={{__html:this.state.selectedEvent.Description}}></p>
            </div>
          </div>
        )}
        {/* {
            this.state.showDialog &&
            <Event
              event={this.state.selectedEvent}
              panelMode={this.state.panelMode}
              onDissmissPanel={this.onDismissPanel}
              showPanel={this.state.showDialog}
              startDate={this.state.startDateSlot}
              endDate={this.state.endDateSlot}
              context={this.props.context}
              siteUrl={this.props.siteUrl}
              listId={this.props.list}
            />
          } */}
      </div>
    );
  }
}
