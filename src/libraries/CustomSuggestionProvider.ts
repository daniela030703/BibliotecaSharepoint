import { BaseSuggestionProvider, ISuggestion } from "@pnp/modern-search-extensibility";
import { IPropertyPaneGroup, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

export interface ICustomSuggestionProviderProperties {
  myProperty: string;
}

export class CustomSuggestionProvider extends BaseSuggestionProvider<ICustomSuggestionProviderProperties> {

  private _zeroTermSuggestions: ISuggestion[] = [];
  
  public async onInit(): Promise<void> {

    this.context.spHttpClient
      .get(`${this.context.pageContext.web.absoluteUrl}/_api/lists/getbytitle('Documentos')/items?$select=FileLeafRef`,
        SPHttpClient.configurations.v1,
        {
          headers: [
            ['accept', 'application/json;odata.metadata=none']
          ]
        })
      .then((res: SPHttpClientResponse): Promise<{ Items: string; }> => {
        return res.json();
      })
      .then((web: any): any => {
        web.value.map((mi_item: any) => {
          const asd = {
            displayText: mi_item.FileLeafRef
          }
          this._zeroTermSuggestions.push(asd); 
        });
      }); 
  }

  public get isZeroTermSuggestionsEnabled(): boolean {
    return true;
  }

  public async getSuggestions(queryText: string): Promise<ISuggestion[]> {
    return await this._getSampleSuggestions(queryText);
  }

  private async getData(): Promise<ISuggestion[]> {
    return this._zeroTermSuggestions;
  }

  public async getZeroTermSuggestions(): Promise<ISuggestion[]> {

    const tmp = await this.getData()
    return tmp
  }

  private _getSampleSuggestions = async (queryText: string): Promise<ISuggestion[]> => {
    const sampleSuggestions = await this.getData()
    return sampleSuggestions.filter(sg => sg.displayText.toLowerCase().match(`\\b${queryText.trim().toLowerCase()}`));
  }

  public getPropertyPaneGroupsConfiguration(): IPropertyPaneGroup[] {

    return [
      {
        groupName: 'Custom Search Suggestions',
        groupFields: [
          PropertyPaneTextField('providerProperties.myProperty', {
            label: 'My property'
          })
        ]
      }
    ];
  }
}