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

export const postToken = async (assertion: string) => {
  const settings = (window as any).appsettings as AppSettings;

  return await axios
    .post(
      `${settings.serverUrl}/api/Microsoft/token`,
      {
        clientId: "",
        grantType: "urn:ietf:params:oauth:grant-type:jwt-bearer",
        assertion: assertion,
        scope: "user.read mail.readwrite",
        requestedTokenUse: "on_behalf_of",
        tenant: "common",
      },
      {
        headers: {
          "X-API-KEY": settings.apiKey,
        },
      }
    )
    .then((res) => {
      return res.data;
    })
    .catch(() => {
      throw new Error("Network request failed");
    });
};

export const PostAttachmentUpload = async (data: FormData) => {
  const settings = (window as any).appsettings as AppSettings;
  return await axios
    .post(`${settings.serverUrl}/api/Attachment/upload`, data, {
      headers: {
        "X-API-KEY": settings.apiKey,
      },
    })
    .then((res) => res.data)
    .catch(() => {
      throw new Error("Network request failed");
    });
};
