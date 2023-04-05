import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { PageContext } from "@microsoft/sp-page-context";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import {  IExtensibilityLibrary, 
          IComponentDefinition, 
          ISuggestionProviderDefinition, 
          ISuggestionProvider,
          ILayoutDefinition, 
          IQueryModifierDefinition,
          IDataSourceDefinition,
          IDataSource
} from "@pnp/modern-search-extensibility";
import * as Handlebars from "handlebars";
import { CustomSuggestionProvider } from "../CustomSuggestionProvider";
import { CustomDataSource } from "../CustomDataSource";
export class LibreriaSugerenciasLibrary implements IExtensibilityLibrary {
 
  

  public static readonly serviceKey: ServiceKey<LibreriaSugerenciasLibrary> =
  ServiceKey.create<LibreriaSugerenciasLibrary>('SPFx:LibreriaSugerencias', LibreriaSugerenciasLibrary);

  private _spHttpClient: SPHttpClient;
  private _pageContext: PageContext;
  private _currentWebUrl: string;

  constructor(serviceScope: ServiceScope) {
    serviceScope.whenFinished(() => {
      this._spHttpClient = serviceScope.consume(SPHttpClient.serviceKey);
      this._pageContext = serviceScope.consume(PageContext.serviceKey);
      this._currentWebUrl = this._pageContext.web.absoluteUrl;
    });
  }

  public getCustomLayouts(): ILayoutDefinition[] {
    return [ ];
  }

  public getCustomWebComponents(): IComponentDefinition<any>[] {
    return [  ];
  }
  public getCustomSuggestionProviders(): ISuggestionProviderDefinition[] {
    return [
        {
          name: 'Proveedor Personalizado de Sugerencias',
          key: 'ProveedorPersonalizadoSugerencias',
          description: 'A demo custom suggestions provider from the extensibility library',
          serviceKey: ServiceKey.create<ISuggestionProvider>('MyCompany:BibliotecaSharepoint', CustomSuggestionProvider)
      }
    ];
  }

  public registerHandlebarsCustomizations?(handlebarsNamespace: typeof Handlebars): void { }

  public invokeCardAction(action: any): void {
    
    // Process the action based on type
    if (action.type == "Action.OpenUrl") {
      window.open(action.url, "_blank");
    } else if (action.type == "Action.Submit") {
      // Process the action based on title
      switch (action.title) {

        case 'Click on item':

           // Invoke the currentUser endpoing
           this._spHttpClient.get(
            `${this._currentWebUrl}/_api/web/currentUser`,
            SPHttpClient.configurations.v1, 
            null).then((response: SPHttpClientResponse) => {
              response.json().then((json) => {
                console.log(JSON.stringify(json));
              });
            });

          break;

        case 'Global click':
          alert(JSON.stringify(action.data));
          break;
        default:
          console.log('Action not supported!');
          break;
      }
    }
  }

  public getCustomQueryModifiers(): IQueryModifierDefinition[]
  {
    return [ ];
  
    }
  public getCustomDataSources(): IDataSourceDefinition[] {
    return [
      {
          name: 'NPM Search',
          iconName: 'Database',
          key: 'CustomDataSource',
          serviceKey: ServiceKey.create<IDataSource>('MyCompany:CustomDataSource', CustomDataSource)
      }
    ];
  }

  public name(): string {
    return 'LibreriaSugerencias';
  }
}
