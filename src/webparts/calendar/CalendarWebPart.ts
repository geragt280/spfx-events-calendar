import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  // PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import { IColumnReturnProperty, PropertyFieldColumnPicker, PropertyFieldColumnPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldColumnPicker';
import * as strings from 'CalendarWebPartStrings';
import Calendar from './components/Calendar';
import { ICalendarProps } from './components/ICalendarProps';
import { spfi, SPFI, SPFx } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/lists/web";
import "@pnp/sp/fields";
import "@pnp/sp/regional-settings/web";
import "@pnp/sp/profiles";

export interface ICalendarWebPartProps {
  headingTitleColor: string;
  title: string;
  feedsTitle: string;
  list: string;
  list2: string;
  errorMessage: string;
  headerColor:string;
  calendarCellColor: string;
  showEventsFeedsWP: boolean;
  list1column1: string;
  list1column2: string;
  list1column3: string;
  list1column4: string;
  list1column5: string;
  list2column1: string;
  list2column2: string;
  list2column3: string;
  list2column4: string;
  list2column5: string;
}

export default class CalendarWebPart extends BaseClientSideWebPart<ICalendarWebPartProps> {

  // private lists: IPropertyPaneDropdownOption[] = [];
  // private lists2: IPropertyPaneDropdownOption[] = [];
  // private listsDropdownDisabled: boolean = true;
  private _sp: SPFI = null;
  // private errorMessage: string;

  public render(): void {
    const element: React.ReactElement<ICalendarProps> = React.createElement(
      Calendar,
      {
        title: this.properties.title,
        feedsTitle: this.properties.feedsTitle,
        list: this.properties.list,
        list1column1: this.properties.list1column1,
        list1column2: this.properties.list1column2,
        list1column3: this.properties.list1column3,
        list1column4: this.properties.list1column4,
        list1column5: this.properties.list1column5,
        list2: this.properties.list2,
        list2column1: this.properties.list2column1,
        list2column2: this.properties.list2column2,
        list2column3: this.properties.list2column3,
        list2column4: this.properties.list2column4,
        list2column5: this.properties.list2column5,
        displayMode: this.displayMode,
        updateTitleProperty: (value: string) => {
          this.properties.title = value
        },
        updateFeedsTitleProperty: (value: string) => {
          this.properties.feedsTitle = value
        },
        context: this.context,
        headerColor: this.properties.headerColor,
        calendarCellColor: this.properties.calendarCellColor,
        themeVariant: null,
        headingTitleColor: this.properties.headingTitleColor,
        showEventsFeedsWP: this.properties.showEventsFeedsWP,
        sp: this._sp
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._sp = spfi().using(SPFx(this.context));
    });
  }
  
  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }
          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: string, newValue: string) {
  //   try {
  //     // reset any error
  //     debugger;
  //     // this.properties.errorMessage = undefined;
  //     // this.errorMessage = undefined;
  //     // this.context.propertyPane.refresh();

  //     switch (propertyPath) {
  //       case "list":
  //         this.properties.list = newValue;
  //         this.context.propertyPane.refresh();
  //         this.render();
  //         break;
  //       case "list2":
  //         this.properties.list2 = newValue;
  //         this.context.propertyPane.refresh();
  //         this.render();
  //         break;      
  //       default:
  //         break;
  //     }
  //   } catch (error) {
  //     this.errorMessage =  `${error.message} -  please check if list is valid.` ;
  //     console.log("Error:", this.errorMessage);
  //     this.context.propertyPane.refresh();
  //   }
  // }

  onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void{
    debugger;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: "Calendar Events Configuration",
              groupFields: [
                PropertyFieldColorPicker('headingTitleColor', {
                  label: 'Webpart Heading background color',
                  selectedColor: this.properties.headingTitleColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  // debounce: 100,
                  isHidden: false,
                  alphaSliderHidden: true,
                  style: PropertyFieldColorPickerStyle.Inline,
                  iconName: 'Precipitation',
                  key: 'headerTitleColorFieldId'
                }),
                // PropertyPaneTextField('siteUrl', {
                //   label: strings.SiteUrlFieldLabel,
                //   onGetErrorMessage: this.onSiteUrlGetErrorMessage.bind(this),
                //   value: this.context.pageContext.site.absoluteUrl,
                //   deferredValidationTime: 1200,
                // }),
                PropertyFieldListPicker('list', {
                  label: 'Select list',
                  selectedList: this.properties.list,
                  includeHidden: false,
                  baseTemplate: [106, 100],
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId'
                }),
                PropertyFieldColumnPicker('list1column1', {
                  label: 'Select title column',
                  context: this.context,
                  selectedColumn: this.properties.list1column1,
                  listId: this.properties.list,
                  disabled: false,
                  orderBy: PropertyFieldColumnPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'columnPickerFieldId1',
                  displayHiddenColumns: false,
                  columnReturnProperty: IColumnReturnProperty["Internal Name"]
                }),
                PropertyFieldColumnPicker('list1column2', {
                  label: 'Select start date column',
                  context: this.context,
                  selectedColumn: this.properties.list1column2,
                  listId: this.properties.list,
                  disabled: false,
                  orderBy: PropertyFieldColumnPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'columnPickerFieldId2',
                  displayHiddenColumns: false,
                  columnReturnProperty: IColumnReturnProperty["Internal Name"]
                }),
                PropertyFieldColumnPicker('list1column3', {
                  label: 'Select end date column',
                  context: this.context,
                  selectedColumn: this.properties.list1column3,
                  listId: this.properties.list,
                  disabled: false,
                  orderBy: PropertyFieldColumnPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'columnPickerFieldId3',
                  displayHiddenColumns: false,
                  columnReturnProperty: IColumnReturnProperty["Internal Name"]
                }),
                PropertyFieldColumnPicker('list1column4', {
                  label: 'Select events back color column',
                  context: this.context,
                  selectedColumn: this.properties.list1column4,
                  listId: this.properties.list,
                  disabled: false,
                  orderBy: PropertyFieldColumnPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'columnPickerFieldId4',
                  displayHiddenColumns: false,
                  columnReturnProperty: IColumnReturnProperty["Internal Name"]
                }),
                PropertyFieldColumnPicker('list1column5', {
                  label: 'Select all day event column',
                  context: this.context,
                  selectedColumn: this.properties.list1column5,
                  listId: this.properties.list,
                  disabled: false,
                  orderBy: PropertyFieldColumnPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'columnPickerFieldId4',
                  displayHiddenColumns: false,
                  columnReturnProperty: IColumnReturnProperty["Internal Name"]
                }),
                // PropertyFieldDateTimePicker('eventStartDate', {
                //   label: 'From',
                //   initialDate: this.properties.eventStartDate,
                //   dateConvention: DateConvention.Date,
                //   onPropertyChange: this.onPropertyPaneFieldChanged,
                //   properties: this.properties,
                //   onGetErrorMessage: this.onEventStartDateValidation,
                //   deferredValidationTime: 0,
                //   key: 'eventStartDateId'
                // }),
                // PropertyFieldDateTimePicker('eventEndDate', {
                //   label: 'to',
                //   initialDate:  this.properties.eventEndDate,
                //   dateConvention: DateConvention.Date,
                //   onPropertyChange: this.onPropertyPaneFieldChanged,
                //   properties: this.properties,
                //   onGetErrorMessage:  this.onEventEndDateValidation,
                //   deferredValidationTime: 0,
                //   key: 'eventEndDateId'
                // }),
                // PropertyFieldColorPicker('headerColor', {
                //   label: 'Calendar header background color',
                //   selectedColor: this.properties.headerColor,
                //   onPropertyChange: this.onPropertyPaneFieldChanged,
                //   properties: this.properties,
                //   disabled: false,
                //   // debounce: 100,
                //   isHidden: false,
                //   alphaSliderHidden: true,
                //   style: PropertyFieldColorPickerStyle.Inline,
                //   iconName: 'Precipitation',
                //   key: 'headerColorFieldId'
                // })
                // PropertyFieldColorPicker('calendarCellColor', {
                //   label: 'Calender cell color',
                //   selectedColor: this.properties.calendarCellColor,
                //   onPropertyChange: this.onPropertyPaneFieldChanged,
                //   properties: this.properties,
                //   disabled: false,
                //   // debounce: 100,
                //   isHidden: false,
                //   alphaSliderHidden: true,
                //   style: PropertyFieldColorPickerStyle.Inline,
                //   iconName: 'Precipitation',
                //   key: 'calendarCellColorFieldId'
                // })
              ]
            }
          ]
        },
        {
          header: {
            description: "Event Feeds Configuration"
          },
          groups: [
            {
              groupName: "Properties",
              groupFields: [
                PropertyPaneToggle('showEventsFeedsWP', {
                  label:'Enable feeds component',
                  checked: this.properties.showEventsFeedsWP
                }),
                PropertyFieldListPicker('list2', {
                  label: 'Select feeds list',
                  selectedList: this.properties.list2,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: !this.properties.showEventsFeedsWP,
                  baseTemplate: [106, 100],
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'list2PickerFieldId'
                }),                
                PropertyFieldColumnPicker('list2column1', {
                  label: 'Select title column',
                  context: this.context,
                  selectedColumn: this.properties.list2column1,
                  listId: this.properties.list2,
                  disabled: !this.properties.showEventsFeedsWP,
                  orderBy: PropertyFieldColumnPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'columnPickerFieldId6',
                  displayHiddenColumns: false,
                  columnReturnProperty: IColumnReturnProperty["Internal Name"]
                }),
                PropertyFieldColumnPicker('list2column2', {
                  label: 'Select start date column',
                  context: this.context,
                  selectedColumn: this.properties.list2column2,
                  listId: this.properties.list2,
                  disabled: !this.properties.showEventsFeedsWP,
                  orderBy: PropertyFieldColumnPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'columnPickerFieldId7',
                  displayHiddenColumns: false,
                  columnReturnProperty: IColumnReturnProperty["Internal Name"]
                }),
                PropertyFieldColumnPicker('list2column3', {
                  label: 'Select end date column',
                  context: this.context,
                  selectedColumn: this.properties.list2column3,
                  listId: this.properties.list2,
                  disabled: !this.properties.showEventsFeedsWP,
                  orderBy: PropertyFieldColumnPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'columnPickerFieldId8',
                  displayHiddenColumns: false,
                  columnReturnProperty: IColumnReturnProperty["Internal Name"]
                }),
                PropertyFieldColumnPicker('list2column4', {
                  label: 'Select events back color column',
                  context: this.context,
                  selectedColumn: this.properties.list2column4,
                  listId: this.properties.list2,
                  disabled: !this.properties.showEventsFeedsWP,
                  orderBy: PropertyFieldColumnPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'columnPickerFieldId9',
                  displayHiddenColumns: false,
                  columnReturnProperty: IColumnReturnProperty["Internal Name"]
                }),
                PropertyFieldColumnPicker('list2column5', {
                  label: 'Select all day event column',
                  context: this.context,
                  selectedColumn: this.properties.list2column5,
                  listId: this.properties.list2,
                  disabled: !this.properties.showEventsFeedsWP,
                  orderBy: PropertyFieldColumnPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'columnPickerFieldId10',
                  displayHiddenColumns: false,
                  columnReturnProperty: IColumnReturnProperty["Internal Name"]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
