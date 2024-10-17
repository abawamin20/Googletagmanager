import { Log } from "@microsoft/sp-core-library";
import { BaseApplicationCustomizer } from "@microsoft/sp-application-base";

declare global {
  interface Window {
    dataLayer: any[];
    gtag: (...args: any[]) => void;
  }
}
/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IGoogleTagManagerApplicationCustomizerProperties {
  googleMeasurementId: string;
  scriptFileName: string;
  scriptFileUrl: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class GoogleTagManagerApplicationCustomizer extends BaseApplicationCustomizer<IGoogleTagManagerApplicationCustomizerProperties> {
  public onInit(): Promise<void> {
    Log.info("GoogleMeasurementId", this.properties.googleMeasurementId);
    Log.info(
      "GoogleTagManagerApplicationCustomizer",
      "GoogleTagManagerApplicationCustomizer loaded"
    );
    const script = document.createElement("script");
    script.src =
      this.properties.scriptFileUrl ||
      `https://www.googletagmanager.com/gtag/js?id=${this.properties.googleMeasurementId}`;
    script.async = true;
    document.body.appendChild(script);

    script.onload = () => {
      window["dataLayer"] = window["dataLayer"] || [];
      window["gtag"] = function () {
        window.dataLayer.push(arguments);
      };
      window["gtag"]("js", new Date());
      window["gtag"]("config", this.properties.googleMeasurementId);
      Log.info("GoogleMeasurementId", this.properties.googleMeasurementId);
      Log.info(
        "GoogleTagManagerApplicationCustomizer",
        "GoogleTagManagerApplicationCustomizer loaded"
      );
    };

    return Promise.resolve();
  }
}
