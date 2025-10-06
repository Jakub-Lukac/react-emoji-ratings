import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneLabel,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ReactEmojiReactionRatingWebPartStrings';
import ReactEmojiReactionRating from './components/ReactEmojiReactionRating';
import { IReactEmojiReactionRatingProps } from './components/IReactEmojiReactionRatingProps';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';

import spService from './components/services/spService';

interface IEmojiItem {
  Title: string;
  ImageUrl: string;
}
// These properties are used to define the configuration of the web part
// and will be set in the property pane of the web part.
export interface IReactEmojiReactionRatingWebPartProps {
  propertyRatingText: string;
  propertyEmojisCollection: IEmojiItem[];
  propertyEnableComments: boolean;
  propertyEnableCount: boolean;
  propertySelectedColor: string;
  propertySelectedEmojiColor: string;
  propertyListName: string;
  propertyCollectionName: string;
  propertyColListColumns: any[];
  propertyListOperationMessage: string;
  propertyLibraryInitialized: boolean;
  propertyCentralSiteUrl: string;
}

export default class ReactEmojiReactionRatingWebPart extends BaseClientSideWebPart<IReactEmojiReactionRatingWebPartProps> {

  private _spService: spService = null;
  public render(): void {

    const element: React.ReactElement<IReactEmojiReactionRatingProps> = React.createElement(
      ReactEmojiReactionRating,
      {
        ratingText: this.properties.propertyRatingText,
        emojisCollection: this.properties.propertyEmojisCollection,
        context: this.context,
        enableComments: this.properties.propertyEnableComments,
        enableCount: this.properties.propertyEnableCount,
        selectedColor: this.properties.propertySelectedColor,
        selectedEmojiColor: this.properties.propertySelectedEmojiColor,
        listName: this.properties.propertyListName,
        collectionName: this.properties.propertyCollectionName,
        displayMode: this.displayMode,
        listMessage: this.properties.propertyListOperationMessage,
        centralSiteUrl: this.properties.propertyCentralSiteUrl,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  public async onInit(): Promise<void> {
    this._spService = new spService(this.context);

    // No need to set propertyListName or propertyColListColumns manually here
    // SPFx will use values from manifest.json

    // Check list initialization status
    // Only attempt to create the list if it hasn't been initialized yet
    if (!this.properties.propertyListOperationMessage) {
      try {

        const centralizedSiteUrl = this.properties.propertyCentralSiteUrl;

        // Create the list if it doesn't exist
        const res = await this._spService.createList(
          this.properties.propertyListName,
          this.properties.propertyColListColumns,
          centralizedSiteUrl
        );

        this.properties.propertyListOperationMessage = "List created successfully";
        this.context.propertyPane.refresh();
        console.log(res);
      } catch (error) {
        const errMessage = error?.message || error?.Message || "Unknown error";
        console.error("List creation failed:", error);
        this.properties.propertyListOperationMessage = `List creation failed: ${errMessage}`;
      }
    }

    return Promise.resolve();
  }

  // This method is used to define the property pane configuration
  // It specifies the fields that will be displayed in the property pane
  // and their labels, types, and other properties.
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName, // "Rating WebPart Configuration"
              groupFields: [  
                PropertyPaneLabel('', {
                  text: strings.ListNameFieldLabel // "List Name"
                }),
                PropertyPaneLabel('propertyListName', {
                  text: this.properties.propertyListName // "CustomerFeedback"
                }),
                PropertyPaneTextField('propertyRatingText', { // propertyRatingText is variable that holds the rating text
                  label: strings.RatingTextFieldLabel // "Rating Text"
                }),
                PropertyPaneLabel('', {
                  text: strings.EmojiLibraryFieldLabel // "Emoji Document Library"
                }),
                PropertyPaneLabel('propertyCollectionName', {
                  text: this.properties.propertyCollectionName // "Emojis Collection"
                }),
                PropertyPaneToggle('propertyEnableComments', {
                  label: strings.EnableCommentsFieldLablel
                }),
                PropertyPaneToggle('propertyEnableCount', {
                  label: strings.EnableCountFieldLablel
                }),
                PropertyFieldColorPicker('propertySelectedColor', {
                  label: 'Select background color',
                  selectedColor: this.properties.propertySelectedColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  debounce: 1000,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'Precipitation',
                  key: 'colorFieldId'
                }),
                PropertyFieldColorPicker('propertySelectedEmojiColor', {
                  label: 'Selected emoji background color',
                  selectedColor: this.properties.propertySelectedEmojiColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  debounce: 1000,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'Precipitation',
                  key: 'colorFieldId'
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
