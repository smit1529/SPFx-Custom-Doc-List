import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown, IPropertyPaneChoiceGroupOption,
  PropertyPaneChoiceGroup,
  PropertyPaneCheckbox,
  PropertyPaneToggle,
  PropertyPaneSlider
} from '@microsoft/sp-webpart-base';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';

import * as strings from 'CustomDocListWebPartStrings';
import CustomDocList from './components/CustomDocList';
import { ICustomDocListProps } from './components/ICustomDocListProps';

export interface ICustomDocListWebPartProps {
  WebPartTitle: string;
  description: string;
  siteAbsoluteUrl: string;
  ListName: string;
  GroupByField: string;
  SortByField: string;
  ColumnsToDisplay: string[];
  ListColumns: any[];
  MaxResultsProp: string;
}

export default class CustomDocListWebPart extends BaseClientSideWebPart<ICustomDocListWebPartProps> {
  private listDropDownOptions: IPropertyPaneDropdownOption[] = [];
  private groupByFiledDropDownOptions: IPropertyPaneDropdownOption[] = [];
  private sortByFiledDropDownOptions: IPropertyPaneDropdownOption[] = [];
  private choiceGroupOptions: any[] = [];

  public async render() {
    const element: React.ReactElement<ICustomDocListProps> = React.createElement(
      CustomDocList,
      {
        WebPartTitle: this.properties.WebPartTitle,
        description: this.properties.description,
        siteAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
        ListName: this.properties.ListName,
        GroupByField: this.properties.GroupByField,
        SortByField: this.properties.SortByField,
        ColumnsToDisplay: this.properties.ColumnsToDisplay,
        ListColumns: this.properties.ListColumns,
        MaxResultsProp: this.properties.MaxResultsProp
      }
    );

    if (this.renderedOnce === false) {
      this.GetListData();

      this.groupByFiledDropDownOptions = [];
      this.groupByFiledDropDownOptions.push({ key: "None", text: "List Not Configured" });

      this.sortByFiledDropDownOptions = [];
      this.sortByFiledDropDownOptions.push({ key: "None", text: "List Not Configured" });

      this.choiceGroupOptions = [];
      this.choiceGroupOptions.push({ key: "None", text: "List Not Configured" });
    }

    if (this.properties.ListName != undefined) {
      var data = await this.GetListColumn();
      this.LoadGroupByFiledDropDownValues(data);
      this.LoadSortByFiledDropDownValues(data);
      this.LoadShowColumnsChoiceValues(data);
    }

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private GetListData(): any {
    var self = this;
    var url = this.context.pageContext.web.absoluteUrl + "/_api/web/lists?$filter=Hidden eq false";

    fetch(url, {
      credentials: 'same-origin',
      headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' }
    })
      .then((res) => res.json())
      .then(
        (result) => {
          var listTitle = result.value.map((items) => { return items.Title; });
          self.LoadDropDownValues(listTitle);
        },
        (error) => {
          console.log(error);
        }
      );
  }

  public GetListColumn(): any {
    var self = this;
    var url = this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('` + this.properties.ListName + `')/Fields?$filter=Hidden eq false and ReadOnlyField eq false`;

    return fetch(url, {
      credentials: 'same-origin',
      headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' }
    })
      .then((res) => res.json())
      .then(
        (result) => {
          this.properties.ListColumns = result.value;
          return result.value;
        },
        (error) => {
          console.log(error);
        }
      );
  }

  private LoadDropDownValues(data): any {
    this.listDropDownOptions = [];
    this.listDropDownOptions.push({ key: "Select", text: "Select" });

    data.map((items) => {
      this.listDropDownOptions.push({ key: items, text: items });
    });
  }

  private LoadGroupByFiledDropDownValues(data): void {
    this.groupByFiledDropDownOptions = [];
    this.groupByFiledDropDownOptions.push({ key: "None", text: "None" });

    data.map((items) => {
      this.groupByFiledDropDownOptions.push({ key: items.Title, text: items.Title });
    });
  }

  private LoadSortByFiledDropDownValues(data): void {
    this.sortByFiledDropDownOptions = [];
    this.sortByFiledDropDownOptions.push({ key: "None", text: "None" });

    data.map((items) => {
      this.sortByFiledDropDownOptions.push({ key: items.Title, text: items.Title });
    });
  }

  private LoadShowColumnsChoiceValues(data): void {
    this.choiceGroupOptions = [];

    data.map((items) => {
      this.choiceGroupOptions.push({ key: items.InternalName, text: items.Title });
    });
  }

  // public LoadLayoutTileValues(data): any {
  //   this.choiceGroupOptions = [];
  //   data.map((items, index) => {
  //     this.choiceGroupOptions.push({ key: items.key, text: items.text, imageSrc: items.imageSrc });
  //   });
  // }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('WebPartTitle', {
                  label: strings.TitleFieldLabel,
                  value: "Web-Part Name"
                }),
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneDropdown('ListName', {
                  label: strings.SelectListFieldLabel,
                  options: this.listDropDownOptions,
                  selectedKey: "Select"
                }),
                PropertyPaneDropdown('GroupByField', {
                  label: strings.GroupByFieldLabel,
                  options: this.groupByFiledDropDownOptions,
                  selectedKey: "None",
                  disabled: true
                }),
                PropertyPaneDropdown('SortByField', {
                  label: strings.SortByFieldLabel,
                  options: this.sortByFiledDropDownOptions,
                  selectedKey: "None"
                }),
                // PropertyPaneToggle('LayoutStyleProp', {
                //   key: 'toBeToggle', label: 'Layout Style',
                //   onText: 'Grid', offText: 'List'
                // }),
                PropertyFieldMultiSelect('ColumnsToDisplay', {
                  key: 'multiSelect',
                  label: strings.ShoeColumnsLabel,
                  options: this.choiceGroupOptions,
                  selectedKeys: this.properties.ColumnsToDisplay
                }),
                PropertyPaneSlider('MaxResultsProp', { label: 'Max results', min: 0, max: 100, step: 5, showValue: true, value: 10 })
              ]
            }
          ]
        }
      ]
    };
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any) {
    // console.log(propertyPath, oldValue, newValue, this.properties);

    if (propertyPath == 'ListName') {
      if (newValue != oldValue) {
        this.properties.ColumnsToDisplay = [];

        var data = await this.GetListColumn();

        this.LoadGroupByFiledDropDownValues(data);
        this.LoadSortByFiledDropDownValues(data);
        this.LoadShowColumnsChoiceValues(data);

        this.context.propertyPane.refresh();
      }
    }
    else if (propertyPath == 'GroupByField') {
      if (newValue != oldValue) {
        // this.context.propertyPane.refresh();
      }
    }
    else if (propertyPath == 'SortByField') {
      if (newValue != oldValue) {
        // this.context.propertyPane.refresh();
      }
    }
    else if (propertyPath == 'ColumnsToDisplay') {
      if (newValue != oldValue) {
        // this.context.propertyPane.refresh();
      }
    }
  }
}
