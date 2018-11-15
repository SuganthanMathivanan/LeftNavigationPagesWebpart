import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  WebPartContext
} from '@microsoft/sp-webpart-base';
export interface ILeftNavigationMenuPagesWebpartProps {
  context: WebPartContext;
  description: string;
}
