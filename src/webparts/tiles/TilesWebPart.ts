import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration
} from '@microsoft/sp-webpart-base';

import * as strings from 'TilesWebPartStrings';
import { ITilesProps } from './components/ITilesProps';
import { ITileInfo, LinkTarget } from './ITileInfo';
import { Tiles } from './components/Tiles';

import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/propertyFields/number';

export interface ITilesWebPartProps {
  collectionData: ITileInfo[];
  tileHeight: number;
  title: string;
  BackgroundColor: string;
  FontColor: string;
}

export default class TilesWebPart extends BaseClientSideWebPart<ITilesWebPartProps> {

  // Just for suppress the tslint validation of dinamically loading of this field by using loadPropertyPaneResources()
  // tslint:disable-next-line: no-any
  private _propertyFieldNumber: any;
  // Just for suppress the tslint validation of dinamically loading of this field by using loadPropertyPaneResources()
  // tslint:disable-next-line: no-any
  private _propertyFieldCollectionData: any;
  // Just for suppress the tslint validation of dinamically loading of this field by using loadPropertyPaneResources()
  // tslint:disable-next-line: no-any

  private _customCollectionFieldType: any;

  public render(): void {
    const element: React.ReactElement<ITilesProps> = React.createElement(
      Tiles,
      {
        title: this.properties.title,
        tileHeight: this.properties.tileHeight,
        BackgroundColor: this.properties.BackgroundColor,
        FontColor: this.properties.FontColor,
        collectionData: this.properties.collectionData,
        displayMode: this.displayMode,
        fUpdateProperty: (value: string) => {
          this.properties.title = value;
        },
        fPropertyPaneOpen: this.context.propertyPane.open
      }
    );

    

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  
  // executes only before property pane is loaded.
  
  protected async loadPropertyPaneResources(): Promise<void> {

    //this._customCollectionFieldType = CustomCollectionFieldType;
    this._propertyFieldCollectionData = PropertyFieldCollectionData;
    //this._propertyFieldNumber = PropertyFieldNumber;

    // import additional controls/components
    
    //const { PropertyFieldNumber } = await import(
      /* webpackChunkName: 'pnp-propcontrols-number' */
      /*'@pnp/spfx-property-controls/lib/propertyFields/number'
    );
    const { PropertyFieldCollectionData, CustomCollectionFieldType } = await import(
      /* webpackChunkName: 'pnp-propcontrols-colldata' */
      /*'@pnp/spfx-property-controls/lib/PropertyFieldCollectionData'
    );

    this.propertyFieldNumber = PropertyFieldNumber;
    this.propertyFieldCollectionData = PropertyFieldCollectionData;
    this.customCollectionFieldTypeString = CustomCollectionFieldType.string;
    this.customCollectionFieldTypeFabricIcon = CustomCollectionFieldType.fabricIcon;
    this.customCollectionFieldTypeDropdown = CustomCollectionFieldType.dropdown;
    this.customCollectionFieldType = CustomCollectionFieldType;*/
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
              groupFields: [
                PropertyFieldCollectionData('collectionData', {
                  key: 'collectionData',
                  label: strings.tilesDataLabel,
                  panelHeader: strings.tilesPanelHeader,
                  // tslint:disable-next-line:max-line-length
                  panelDescription: `${strings.iconInformation} https://developer.microsoft.com/en-us/fabric#/styles/icons`,
                  manageBtnLabel: strings.tilesManageBtn,
                  value: this.properties.collectionData,
                  fields: [
                    {
                      id: 'title',
                      title: strings.titleField,
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: 'description',
                      title: strings.descriptionField,
                      type: CustomCollectionFieldType.string,
                      required: false
                    },
                    {
                      id: 'url',
                      title: strings.urlField,
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: 'icon',
                      title: strings.iconField,
                      type: CustomCollectionFieldType.fabricIcon,
                      required: true
                    },
                    {
                      id: 'target',
                      title: strings.targetField,
                      type: CustomCollectionFieldType.dropdown,
                      options: [
                        {
                          key: LinkTarget.parent,
                          text: strings.targetCurrent
                        },
                        {
                          key: LinkTarget.blank,
                          text: strings.targetNew
                        }
                      ]
                    }
                  ]
                }),
                PropertyFieldNumber('tileHeight', {
                  key: 'tileHeight',
                  label: strings.TileHeight,
                  value: this.properties.tileHeight
                })
                
              ]
            },
            {
              groupName : strings.StylingGroup,
              groupFields: [
                PropertyFieldColorPicker('backgroundcolor',{
                  label : strings.BackgroundColor,
                  selectedColor: this.properties.BackgroundColor,
                  onPropertyChange : this.onPropertyPaneFieldChanged.bind(this),
                  properties : this.properties,
                  disabled : false,
                  alphaSliderHidden :false,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'backgroundcolor'
                }),
                PropertyFieldColorPicker('fontColor', {
                  label: strings.FontColor,
                  selectedColor: this.properties.FontColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  disabled: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'fontColor'
                })

              ]
            }
          ]
        }
      ]
    };
  }
}
