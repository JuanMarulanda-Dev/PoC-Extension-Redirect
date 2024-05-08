import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';


/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IPaginaEnBlancoApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class PaginaEnBlancoApplicationCustomizer
  extends BaseApplicationCustomizer<IPaginaEnBlancoApplicationCustomizerProperties> {

    protected onInit(): Promise<void> {
      // Obtiene la URL actual de la página
      const currentUrl = window.location.href;
    
      // Verifica si la URL corresponde a la página específica
      const isSpecificPage = (currentUrl: string) => currentUrl.toLowerCase().indexOf('TestApp.aspx'.toLowerCase()) !== -1;

      // Verifica si la URL ya contiene los parámetros que se van a agregar
      const hasParameters = (currentUrl: string) => currentUrl.toLowerCase().indexOf('Env=Embedded&env=WebViewList'.toLowerCase()) !== -1;

      // Verifica si tiene o no más parametros
      const hasQuestionMark =  (currentUrl: string) => currentUrl.indexOf('?') !== -1;

      // Agrega los parámetros a la URL
      const addParametersToUrl = (currentUrl: string) => {
        return `${currentUrl}${hasQuestionMark(currentUrl) ? '&' : '?'}Env=Embedded&env=WebViewList`;
      };

      if (isSpecificPage(currentUrl) && !hasParameters(currentUrl)) {
        const newUrl = addParametersToUrl(currentUrl);
        window.location.href = newUrl;
      }
    
      return Promise.resolve();
    }
}
