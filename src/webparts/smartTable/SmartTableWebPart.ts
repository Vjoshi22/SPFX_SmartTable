import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
} from "@microsoft/sp-property-pane";
import {
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy,
} from "@pnp/spfx-property-controls/lib/PropertyFieldListPicker";
import { PropertyFieldMultiSelect } from "@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect";
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import {
  BaseClientSideWebPart,
  WebPartContext,
} from "@microsoft/sp-webpart-base";

import * as strings from "SmartTableWebPartStrings";
import SmartTable from "./components/SmartTable";
import { ISmartTableProps } from "./components/ISmartTableProps";
import { update, get } from "@microsoft/sp-lodash-subset";
//import collection data
import {
  PropertyFieldCollectionData,
  CustomCollectionFieldType,
} from "@pnp/spfx-property-controls/lib/PropertyFieldCollectionData";
import { Web } from "sp-pnp-js";
import { IPropertFieldMultiSelect } from "./components/IPropertyFieldMultiSelect";

var optionVal: Array<IPropertFieldMultiSelect> = new Array<IPropertFieldMultiSelect>();
export interface ISmartTableWebPartProps {
  description: string;
  context: WebPartContext;
  lists: string; // Stores the list ID
  tableTitle: string;
  columnData: any[];
  fieldListCollection: string[];
  selectedKeys: any[];
  color:string;

}

export default class SmartTableWebPart extends BaseClientSideWebPart<
  ISmartTableWebPartProps
> {
  private fieldsListCollection = [];
  public onInit<T>(): Promise<T> {
    this.getFieldsForSelectedList().then((response) => {
      const fields = response;
      optionVal.splice(0,optionVal.length);
      for (let f = 0; f < fields.length; f++) {
        optionVal.push({
          key:fields[f].Title,
          text:fields[f].Title
        })
      }
      this.fieldsListCollection = optionVal;
      return this.fieldsListCollection
      this.context.propertyPane.refresh();
    });
    return Promise.resolve();
  }
  
  public render(): void {
    const element: React.ReactElement<ISmartTableProps> = React.createElement(
      SmartTable,
      {
        description: this.properties.description,
        context: this.context,
        lists: this.properties.lists,
        columnData: this.properties.columnData,
        tableTitle: this.properties.tableTitle,
        fieldListCollection: this.properties.fieldListCollection,
        selectedKeys: this.properties.selectedKeys,
        color: this.properties.color
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

    if(!this.properties.fieldListCollection){
      this.getFieldsForSelectedList().then((response) => {
        const fields = response;
        optionVal.splice(0,optionVal.length);
        for (let f = 0; f < fields.length; f++) {
          optionVal.push({
            key:fields[f].Title,
            text:fields[f].Title
          })
        }
        this.fieldsListCollection = optionVal;
        return this.fieldsListCollection
        this.context.propertyPane.refresh();
      });
    }
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("tableTitle", {
                  label: "Table Title",
                }),
                PropertyFieldListPicker("lists", {
                  label: "Select a list",
                  selectedList: this.properties.lists,
                  baseTemplate: 100,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: "listPickerFieldId",
                }),
                PropertyFieldMultiSelect("fieldListCollection", {
                  key: "fieldListCollection",
                  label: "Select Column to get filtered data",
                  options: this.fieldsListCollection,
                  selectedKeys: this.properties.fieldListCollection
                }),
                PropertyFieldCollectionData("columnData", {
                  key: "columnData",
                  label: "Table Columns",
                  panelHeader:
                    "Enter the Internal Name of columns and Display Name of the selected list",
                  manageBtnLabel: "Click to add columns",
                  value: this.properties.columnData,
                  fields: [
                    {
                      id: "internalName",
                      title: "Internal Name of Column",
                      type: CustomCollectionFieldType.string,
                      required: true,
                    },
                    {
                      id: "displayName",
                      title: "Display Name of Column",
                      type: CustomCollectionFieldType.string,
                    },
                  ],
                  disabled: false,
                }),
              ],
            }
          ],
        },
        {
          header: {
            description: 'Other Configurations',
          },
          groups: [
            {
              groupName: 'Page 2',
              groupFields: [
                PropertyFieldColorPicker('color', {
                  label: 'Color',
                  selectedColor: this.properties.color,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'Precipitation',
                  key: 'colorFieldId'
                })
              ]
            }
          ],
        }
      ],
    };
  }
  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: any,
    newValue: any
  ): void {
    if (propertyPath === "lists" && newValue) {
      this.properties.lists = newValue;
      this.properties.fieldListCollection = [];
      this.getFieldsForSelectedList().then((response) => {
        const fields = response;
        optionVal.splice(0,optionVal.length);
        for (let f = 0; f < fields.length; f++) {
          optionVal.push({
            key:fields[f].InternalName,
            text:fields[f].InternalName
          })
        }
        this.fieldsListCollection = optionVal;
        return optionVal;
        // this.context.propertyPane.refresh();
      });
    } else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    }
  }
  private OnListChange(propertyPath: string, newValue: any): void {
    const oldValue: any = get(this.properties, propertyPath);
    if (oldValue !== newValue) {
      this.properties.lists = null;
    }
    // store new value in web part properties
    update(this.properties, propertyPath, (): any => newValue);
    // refresh property Pane
    this.context.propertyPane.refresh();
    // refresh web part
    this.render();
  }
  private getFieldsForSelectedList = async (): Promise<any> => {
    if (!this.properties.lists) {
      return Promise.resolve();
    }
    const filter = "Hidden eq false and ReadOnlyField eq false";
    let web = new Web(this.context.pageContext.web.absoluteUrl);
    const data = await web.lists
      .getById(this.properties.lists)
      .fields.select("InternalName")
      .filter(filter)
      .get();
    return data;
  };
}
