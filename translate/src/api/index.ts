import { AppSettings } from "../../appsettings";

// /api/Google/translate data: any  JSON.stringify(data)
export const PostTranslate = (content: string) => {
  const settings = (window as any).appsettings as AppSettings;
  return fetch(`${settings.serverUrl}/api/Google/translate`, {
    method: "post",
    body: JSON.stringify({
      content: content,
      targetLanguage: "zh-TW",
    }),
    headers: {
      "X-API-KEY": settings.apiKey,
    },
  })
    .then((res) => res.json())
    .then((res) => {
      return res;
    })
    .catch((err) => console.log(err, "err"));
};
