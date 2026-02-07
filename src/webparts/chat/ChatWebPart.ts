import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IChatWebPartProps } from './IChatWebPartProps';
import Chat from './components/Chat';

export default class ChatWebPart extends BaseClientSideWebPart<IChatWebPartProps> {
  public render(): void {
    const element: React.ReactElement = React.createElement(Chat, {
      apiKey: this.properties.apiKey,
      context: this.context,
      docsLibraryName: this.properties.docsLibraryName ?? '',
      isDarkTheme: !!((this.context.sdks as { theme?: { isInverted?: boolean } })?.theme?.isInverted),
      hasTeamsContext: !!((this.context.sdks as { teams?: unknown })?.teams),
      userDisplayName: this.context.pageContext?.user?.displayName ?? ''
    });
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
            description: 'Configure the Chat web part. Add your OpenAI API key here, or store it in a SharePoint list and load it in the app.'
          },
          groups: [
            {
              groupName: 'OpenAI',
              groupFields: [
                PropertyPaneTextField('apiKey', {
                  label: 'OpenAI API key',
                  description: 'Optional. You can also provide the key from a config list in the app.',
                  multiline: false
                })
              ]
            },
            {
              groupName: 'Storage',
              groupFields: [
                PropertyPaneTextField('docsLibraryName', {
                  label: 'Chats document library',
                  description: 'Document library on this site to store chats (filtered by user). Leave blank to use browser storage only.',
                  multiline: false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
