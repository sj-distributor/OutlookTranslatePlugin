import { AppSettings } from "../../appsettings";
import axios from "axios";

export const PostTranslate = (content: string, language: string) => {
  const settings = (window as any).appsettings as AppSettings;
  return axios
    .post(
      `${settings.serverUrl}/api/Google/translate`,
      {
        isHtml: true,
        content: content,
        targetLanguage: language,
      },
      {
        headers: {
          "X-API-KEY": settings.apiKey,
        },
      }
    )
    .then((response) => {
      return response.data as string;
    });
};
