export interface AppSettings {
  serverUrl: string;
  token: string;
}

const settings = (window as any).appsettings;

export async function InitialAppSetting() {
  if ((window as any).appsettings) return (window as any).appsettings;

  const appSettingsData = require("./appsetting.json");

  (window as any).appsettings = appSettingsData;
}

export default settings as AppSettings;
