import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  WebPartContext
} from '@microsoft/sp-webpart-base';
import * as strings from 'LeftNavigationMenuPagesWebpartWebPartStrings';
import DisplayDocuments from './components/LeftNavigationMenuPagesWebpart';
import MatterDocuments from './components/MatterDocuments';
import MyDocuments from './components/MyDocuments';
import ContactDocuments from './components/ContactDocuments';
import OppoDocuments from './components/OpportunitiesDocs';
import OtherDocuments from './components/Others';
import * as pnp from 'sp-pnp-js';
import { SPComponentLoader } from '@microsoft/sp-loader';
//import { IDisplayDocumentsProps } from './components/IDisplayDocumentsProps';

export interface IDisplayDocumentsWebPartProps {
  context: WebPartContext;
  description: string;
  
}

export default class DisplayDocumentsWebPart extends BaseClientSideWebPart<IDisplayDocumentsWebPartProps> {
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      SPComponentLoader.loadCss('https://use.fontawesome.com/releases/v5.3.1/css/all.css');
      pnp.setup({
        spfxContext: this.context
      });
      
    });
  }
  public render(): void {
    if(this.properties.description == 'Client'){
      const element: React.ReactElement<IDisplayDocumentsWebPartProps > = React.createElement(
        DisplayDocuments,
        {
          description: this.properties.description,
          context: this.context
        }
      );
      ReactDom.render(element, this.domElement);
    }
    if(this.properties.description == 'Matter'){
      const element: React.ReactElement<IDisplayDocumentsWebPartProps > = React.createElement(
        MatterDocuments,
        {
          description: this.properties.description,
          context: this.context
        }
      );
      ReactDom.render(element, this.domElement);
    }
    if(this.properties.description == 'MyDocuments'){
      const element: React.ReactElement<IDisplayDocumentsWebPartProps > = React.createElement(
        MyDocuments,
        {
          description: this.properties.description,
          context: this.context
        }
      );
      ReactDom.render(element, this.domElement);
    }
    if(this.properties.description == 'Contact'){
      const element: React.ReactElement<IDisplayDocumentsWebPartProps > = React.createElement(
        ContactDocuments,
        {
          description: this.properties.description,
          context: this.context
        }
      );
      ReactDom.render(element, this.domElement);
    }
    if(this.properties.description == 'Opportunities'){
      const element: React.ReactElement<IDisplayDocumentsWebPartProps > = React.createElement(
        OppoDocuments,
        {
          description: this.properties.description,
          context: this.context
        }
      );
      ReactDom.render(element, this.domElement);
    }
    if(this.properties.description == 'Others'){
      const element: React.ReactElement<IDisplayDocumentsWebPartProps > = React.createElement(
        OtherDocuments,
        {
          description: this.properties.description,
          context: this.context
        }
      );
      ReactDom.render(element, this.domElement);
    }

    //ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
