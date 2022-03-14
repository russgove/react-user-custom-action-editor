import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'UserCustomActionEditorWebPartStrings';
import UserCustomActionEditor from './components/UserCustomActionEditor';
import { IUserCustomActionEditorProps } from './components/IUserCustomActionEditorProps';

export interface IUserCustomActionEditorWebPartProps {
  description: string;
}
import { spfi, SPFx } from "@pnp/sp";
export default class UserCustomActionEditorWebPart extends BaseClientSideWebPart<IUserCustomActionEditorWebPartProps> {
  protected async onInit(): Promise<void> {

    await super.onInit();
    const sp = spfi().using(SPFx(this.context));

}
  public render(): void {
    debugger;
    const element: React.ReactElement<IUserCustomActionEditorProps> = React.createElement(
      UserCustomActionEditor,
      {
        description: this.properties.description,
        context:this.context
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
