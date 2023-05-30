export interface AppSettings {
  serverUrl: string;
  apiKey: string;
}

const settings = (window as any).appsettings;

export async function InitialAppSetting() {
  if ((window as any).appsettings) return (window as any).appsettings;

  // const appSettingsData = require("./appsetting.json");

  fetch("./appsetting.json")
    .then((response) => response.json())
    .then((data: AppSettings) => {
      (window as any).appsettings = data;
    });
}

export default settings as AppSettings;
