import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import * as strings from 'CalendarWebPartStrings';
import Calendar from './components/Calendar';
import { ICalendarProps } from './components/ICalendarProps';
import { PropertyFieldDateTimePicker, DateConvention, TimeConvention, IDateTimeFieldValue } from '@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker';
import { ThemeProvider, ThemeChangedEventArgs, IReadonlyTheme, ISemanticColors } from '@microsoft/sp-component-base';
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
// import { PropertyPaneHorizontalRule } from "@microsoft/sp-property-pane";

export interface ICalendarWebPartProps {
  headingTitleColor: string;
  title: string;
  feedsTitle: string;
  siteUrl: string;
  list: string;
  eventStartDate: IDateTimeFieldValue ;
  eventEndDate: IDateTimeFieldValue;
  siteUrl2: string;
  list2: string;
  eventStartDate2: IDateTimeFieldValue ;
  eventEndDate2: IDateTimeFieldValue;
  errorMessage: string;
  headerColor:string;
  calendarCellColor: string;
}
import spservices from '../../services/spservices';
import * as moment from 'moment';

export default class CalendarWebPart extends BaseClientSideWebPart<ICalendarWebPartProps> {

  private lists: IPropertyPaneDropdownOption[] = [];
  private lists2: IPropertyPaneDropdownOption[] = [];
  private listsDropdownDisabled: boolean = true;
  private spService: spservices = null;
  private errorMessage: string;
  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;

  public constructor() {
    super();

  }

  public render(): void {

    const element: React.ReactElement<ICalendarProps> = React.createElement(
      Calendar,
      {
        title: this.properties.title,
        feedsTitle: this.properties.feedsTitle,
        siteUrl: this.properties.siteUrl,
        list: this.properties.list,
        eventStartDate: this.properties.eventStartDate,
        eventEndDate: this.properties.eventEndDate,
        siteUrl2: this.properties.siteUrl2,
        list2: this.properties.list2,
        eventStartDate2: this.properties.eventStartDate2,
        eventEndDate2: this.properties.eventEndDate2,
        displayMode: this.displayMode,
        headerColor:this.properties.headerColor,
        calendarCellColor: this.properties.calendarCellColor,
        context: this.context,
        updateTitleProperty: (value: string) => {
          this.properties.title = value;
        },
        updateFeedsTitleProperty: (value: string) => {
          this.properties.feedsTitle = value;
        },
        headingTitleColor: this.properties.headingTitleColor
      }
    );

    ReactDom.render(element, this.domElement);
  }

  // onInit
  public  async onInit(): Promise<void> {

    this.spService = new spservices(this.context);
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
    
    this._themeVariant = this._themeProvider.tryGetTheme();
    
    this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);

    this.properties.siteUrl = this.properties.siteUrl ? this.properties.siteUrl : this.context.pageContext.web.absoluteUrl;
    this.properties.siteUrl2 = this.properties.siteUrl2 ? this.properties.siteUrl2 : this.context.pageContext.web.absoluteUrl;
    if (!this.properties.eventStartDate){
      this.properties.eventStartDate = { value: moment().subtract(2,'years').startOf('month').toDate(), displayValue: moment().format('ddd MMM MM YYYY')};
    }
    if (!this.properties.eventEndDate){
      this.properties.eventEndDate = { value: moment().add(20,'years').endOf('month').toDate(), displayValue: moment().format('ddd MMM MM YYYY')};
    }
    if (!this.properties.eventStartDate2){
      this.properties.eventStartDate2 = { value: moment().subtract(2,'years').startOf('month').toDate(), displayValue: moment().format('ddd MMM MM YYYY')};
    }
    if (!this.properties.eventEndDate2){
      this.properties.eventEndDate2 = { value: moment().add(20,'years').endOf('month').toDate(), displayValue: moment().format('ddd MMM MM YYYY')};
    }
    if (this.properties.siteUrl && !this.properties.list) {
     const _lists = await this.loadLists();
     if ( _lists.length > 0 ){
      this.lists = _lists;
      this.properties.list = this.lists[0].key.toString();
     }
    }
    if (this.properties.siteUrl2 && !this.properties.list2) {
      const _lists = await this.loadLists2();
      if ( _lists.length > 0 ){
       this.lists2 = _lists;
       this.properties.list2 = this.lists2[0].key.toString();
      }
     }

    return Promise.resolve();
  }


  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /**
   *
   * @protected
   * @memberof CalendarWebPart
   */
  protected async onPropertyPaneConfigurationStart() {

    try {
      if (this.properties.siteUrl) {
        const _lists = await this.loadLists();
        this.lists = _lists;
        this.listsDropdownDisabled = false;
        //  await this.loadFields(this.properties.siteUrl);
        this.context.propertyPane.refresh();

      } else {
        this.lists = [];
        this.properties.list = '';
        this.listsDropdownDisabled = false;
        this.context.propertyPane.refresh();
      }
      if (this.properties.siteUrl2) {
        const _lists = await this.loadLists2();
        this.lists2 = _lists;
        this.listsDropdownDisabled = false;
        //  await this.loadFields(this.properties.siteUrl);
        this.context.propertyPane.refresh();

      } else {
        this.lists2 = [];
        this.properties.list2 = '';
        this.listsDropdownDisabled = false;
        this.context.propertyPane.refresh();
      }

    } catch (error) {

    }
  }

  /**
   *
   * @private
   * @returns {Promise<IPropertyPaneDropdownOption[]>}
   * @memberof CalendarWebPart
   */
  private async loadLists(): Promise<IPropertyPaneDropdownOption[]> {
    const _lists: IPropertyPaneDropdownOption[] = [];
    try {
      const results = await this.spService.getSiteLists(this.properties.siteUrl);
      for (const list of results) {
        _lists.push({ key: list.Id, text: list.Title });
      }
      // push new item value
    } catch (error) {
      this.errorMessage =  `${error.message} -  please check if site url if valid.` ;
      this.context.propertyPane.refresh();
    }
    return _lists;
  }

  private async loadLists2(): Promise<IPropertyPaneDropdownOption[]> {
    const _lists: IPropertyPaneDropdownOption[] = [];
    try {
      const results = await this.spService.getSiteLists(this.properties.siteUrl2);
      for (const list of results) {
        _lists.push({ key: list.Id, text: list.Title });
      }
      // push new item value
    } catch (error) {
      this.errorMessage =  `${error.message} -  please check if site url if valid.` ;
      this.context.propertyPane.refresh();
    }
    return _lists;
  }

  /**
   *
   *
   * @private
   * @param {string} date
   * @returns
   * @memberof CalendarWebPart
   */
  private onEventStartDateValidation(date:string){
    if (date && this.properties.eventEndDate.value){
      if (moment(date).isAfter(moment(this.properties.eventEndDate.value))){
        return strings.SartDateValidationMessage;
      }
    }
    return '';
  }

  /**
   *
   * @private
   * @param {string} date
   * @returns
   * @memberof CalendarWebPart
   */
  private onEventEndDateValidation(date:string){
    if (date && this.properties.eventEndDate.value){
      if (moment(date).isBefore( moment(this.properties.eventStartDate.value))){
        return strings.EnDateValidationMessage;
      }
    }
    return '';
  }
  /**
   *
   * @private
   * @param {string} value
   * @returns {Promise<string>}
   * @memberof CalendarWebPart
   */

  private onSiteUrlGetErrorMessage(value: string) {
    let returnValue: string = '';
    if (value) {
      returnValue = '';
    } else {
      const previousList: string = this.properties.list;
      const previousSiteUrl: string = this.properties.siteUrl;
      const previousList2: string = this.properties.list2;
      const previousSiteUrl2: string = this.properties.siteUrl2;
      // reset selected item
      this.properties.list = undefined;
      this.properties.siteUrl = undefined;
      this.properties.list2 = undefined;
      this.properties.siteUrl2 = undefined;
      this.lists = [];
      this.lists2 = [];
      this.listsDropdownDisabled = true;
      this.onPropertyPaneFieldChanged('list', previousList, this.properties.list);
      this.onPropertyPaneFieldChanged('siteUrl', previousSiteUrl, this.properties.siteUrl);
      this.onPropertyPaneFieldChanged('list2', previousList2, this.properties.list2);
      this.onPropertyPaneFieldChanged('siteUrl2', previousSiteUrl2, this.properties.siteUrl2);
      this.context.propertyPane.refresh();
    }
    return returnValue;
  }

  /**
   *
   * @protected
   * @param {string} propertyPath
   * @param {string} oldValue
   * @param {string} newValue
   * @memberof CalendarWebPart
   */
  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: string, newValue: string) {
    try {
      // reset any error
      this.properties.errorMessage = undefined;
      this.errorMessage = undefined;
      this.context.propertyPane.refresh();

      if (propertyPath === 'siteUrl' && newValue) {
        super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
        const _oldValue = this.properties.list;
        this.onPropertyPaneFieldChanged('list', _oldValue, this.properties.list);
        this.context.propertyPane.refresh();
        const _lists = await this.loadLists();
        this.lists = _lists;
        this.listsDropdownDisabled = false;
        this.properties.list = this.lists.length > 0 ? this.lists[0].key.toString() : undefined;
        this.context.propertyPane.refresh();
        this.render();
      }else if (propertyPath === 'siteUrl2' && newValue) {
        super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
        const _oldValue = this.properties.list2;
        this.onPropertyPaneFieldChanged('list', _oldValue, this.properties.list2);
        this.context.propertyPane.refresh();
        const _lists = await this.loadLists2();
        this.lists2 = _lists;
        this.listsDropdownDisabled = false;
        this.properties.list2 = this.lists2.length > 0 ? this.lists2[0].key.toString() : undefined;
        this.context.propertyPane.refresh();
        this.render();
      }
      else {
        super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      }
    } catch (error) {
      this.errorMessage =  `${error.message} -  please check if site url if valid.` ;
      this.context.propertyPane.refresh();
    }
  }
  /**
   *
   * @protected
   * @returns {IPropertyPaneConfiguration}
   * @memberof CalendarWebPart
   */
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
      // EndDate and Start Date defualt values

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
                PropertyPaneDropdown('list', {
                  label: strings.ListFieldLabel,
                  options: this.lists,
                  disabled: this.listsDropdownDisabled,
                }),
                // PropertyPaneLabel('eventStartDate', {
                //   text: strings.eventSelectDatesLabel
                // }),
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
                // PropertyPaneLabel('errorMessage', {
                //   text:  this.errorMessage,
                // }),
                PropertyFieldColorPicker('headerColor', {
                  label: 'Calendar header background color',
                  selectedColor: this.properties.headerColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  // debounce: 100,
                  isHidden: false,
                  alphaSliderHidden: true,
                  style: PropertyFieldColorPickerStyle.Inline,
                  iconName: 'Precipitation',
                  key: 'headerColorFieldId'
                }),
                PropertyFieldColorPicker('calendarCellColor', {
                  label: 'Calender cell color',
                  selectedColor: this.properties.calendarCellColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  // debounce: 100,
                  isHidden: false,
                  alphaSliderHidden: true,
                  style: PropertyFieldColorPickerStyle.Inline,
                  iconName: 'Precipitation',
                  key: 'calendarCellColorFieldId'
                })
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
                // PropertyPaneTextField('siteUrl2', {
                //   label: strings.SiteUrlFieldLabel,
                //   onGetErrorMessage: this.onSiteUrlGetErrorMessage.bind(this),
                //   value: this.context.pageContext.site.absoluteUrl,
                //   deferredValidationTime: 1200,
                // }),
                PropertyPaneDropdown('list2', {
                  label: strings.ListFieldLabel,
                  options: this.lists2,
                  disabled: this.listsDropdownDisabled,
                }),
                // PropertyPaneLabel('eventStartDate2', {
                //   text: strings.eventSelectDatesLabel
                // }),
                // PropertyFieldDateTimePicker('eventStartDate2', {
                //   label: 'From',
                //   initialDate: this.properties.eventStartDate2,
                //   dateConvention: DateConvention.Date,
                //   onPropertyChange: this.onPropertyPaneFieldChanged,
                //   properties: this.properties,
                //   onGetErrorMessage: this.onEventStartDateValidation,
                //   deferredValidationTime: 0,
                //   key: 'eventStartDateId2'
                // }),
                // PropertyFieldDateTimePicker('eventEndDate2', {
                //   label: 'to',
                //   initialDate:  this.properties.eventEndDate2,
                //   dateConvention: DateConvention.Date,
                //   onPropertyChange: this.onPropertyPaneFieldChanged,
                //   properties: this.properties,
                //   onGetErrorMessage:  this.onEventEndDateValidation,
                //   deferredValidationTime: 0,
                //   key: 'eventEndDateId2'
                // }),
                // PropertyPaneLabel('errorMessage2', {
                //   text:  this.errorMessage,
                // }),
                // PropertyFieldColorPicker('headerColor', {
                //   label: 'Header background color',
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
                // }),
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
        }
      ]
    };
  }

   /**
 * Update the current theme variant reference and re-render.
 *
 * @param args The new theme
 */
    private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
      this._themeVariant = args.theme;
      this.render();
    }
}
