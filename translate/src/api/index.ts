import { AppSettings } from "../../appsettings";

// /api/Google/translate data: any  JSON.stringify(data)
export const PostTranslate = () => {
  const settings = (window as any).appsettings as AppSettings;
  return fetch(`${settings.serverUrl}/api/MetaShower/list`, {
    method: "get",
    // body: JSON.stringify({
    //   page
    // }),
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
