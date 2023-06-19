import { AppSettings } from "../../appsettings";
import axios from "axios";

export const PostTranslate = async (content: string, language: string, boolean: boolean = true) => {
  const settings = (window as any).appsettings as AppSettings;
  try {
    const response = await axios.post(
      `${settings.serverUrl}/api/Google/translate`,
      {
        isHtml: boolean,
        content: content,
        targetLanguage: language,
      },
      {
        headers: {
          "X-API-KEY": settings.apiKey,
        },
      }
    );
    return response.data;
  } catch (error) {
    throw new Error(error);
  }
};
