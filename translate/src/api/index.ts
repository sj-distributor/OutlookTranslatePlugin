import { AppSettings } from "../../appsettings";
import axios from "axios";

// /api/Google/translate data: any  JSON.stringify(data)
export const PostTranslate = (content: string) => {
  const settings = (window as any).appsettings as AppSettings;
  return axios
    .post(
      `${settings.serverUrl}/api/Google/translate`,
      {
        content: content,
        targetLanguage: "zh-TW",
      },
      {
        headers: {
          "X-API-KEY": settings.apiKey,
        },
      }
    )
    .then(function (response) {
      console.log(response);
    })
    .catch(function (error) {
      console.log(error);
    });
  // return fetch(`${settings.serverUrl}/api/Google/translate`, {
  //   method: "post",
  //   body: {
  //     content: content,
  //     targetLanguage: "zh-TW",
  //   },
  //   headers: {
  //     "X-API-KEY": settings.apiKey,
  //   },
  // })
  //   .then((res) => res.json())
  //   .then((res) => {
  //     return res;
  //   })
  //   .catch((err) => console.log(err, "err"));
};
