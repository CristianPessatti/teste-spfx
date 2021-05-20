import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ListagemMemorandosWebPartStrings';
import ListagemMemorandos from './components/ListagemMemorandos';
import { IListagemMemorandosProps } from './components/IListagemMemorandosProps';

export interface IListagemMemorandosWebPartProps {
  description: string;
}

export default class ListagemMemorandosWebPart extends BaseClientSideWebPart<IListagemMemorandosWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IListagemMemorandosProps> = React.createElement(
      ListagemMemorandos,
      {
        description: this.properties.description,
        siteURL: window.location.origin,
        context: this.context
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
