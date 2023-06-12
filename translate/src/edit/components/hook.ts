/* eslint-disable no-undef */
import { message } from "antd";
import axios from "axios";
import { useEffect, useState } from "react";
import { PublicClientApplication } from "@azure/msal-browser";

export const useAction = () => {
  const [content, setContent] = useState<string>("");

  const [type, setType] = useState<string>("html");

  const [list, setList] = useState<string[]>([]);

  const [src, setSrc] = useState<string>("");

  const [s, setS] = useState<string[]>([]);

  const [d, setD] = useState<
    {
      name: string;
      src: string;
    }[]
  >([]);

  const [z, setZ] = useState<string[]>([]);

  // useEffect(() => {
  //   // 获取类型
  //   // Office.context.mailbox.item.body.getTypeAsync((asyncResult) => {
  //   //   if (asyncResult.status === Office.AsyncResultStatus.Failed) {
  //   //     console.log("Action failed with error: " + asyncResult.error.message);
  //   //     return;
  //   //   }
  //   //   setType(asyncResult.value.toString());
  //   // });
  //   //获取正文
  //   Office.context.mailbox.item.body.getAsync("html", function callback(html) {
  //     setContent(html.value);

  //     const tempDiv = document.createElement("div");
  //     tempDiv.innerHTML = html.value;

  //     const imgTags = tempDiv.getElementsByTagName("img");

  //     const promises = [];

  //     for (let i = 0; i < imgTags.length; i++) {
  //       const img = imgTags[i];
  //       const src = img.getAttribute("src");
  //       promises.push(src);
  //     }

  //     setS(promises);
  //   });

  //   // 获取原始邮件附件信息
  //   // Office.context.mailbox.item.getItemIdAsync((res) => {
  //   //   console.log(res.value);
  //   //   Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, async function (result) {
  //   //     console.log(result.value);
  //   //     var url = `https://outlook.office.com/api/v2.0/me/messages/${res.value}/attachments`;

  //   //     var messageUrl = `https://outlook.office.com/api/v2.0/me/messages/${res.value}`;

  //   //     // fetch(url, {
  //   //     //   headers: {
  //   //     //     Authorization: "Bearer " + result.value,
  //   //     //   },
  //   //     // })
  //   //     //   .then((response) => response.json())
  //   //     //   .then((data) => {
  //   //     //     console.log(data);
  //   //     //     // const list = data.value.map((item) => item["Id"]);
  //   //     //     // setList(list);
  //   //     //     // list.map((item) => {
  //   //     //     //   console.log(item);
  //   //     //     //   Office.context.mailbox.item.getAttachmentContentAsync(item, (result) => {
  //   //     //     //     console.log(result);
  //   //     //     //     setSrc("data:image/png;base64," + result.value.content);
  //   //     //     //   });
  //   //     //     // });
  //   //     //   })
  //   //     //   .catch((error) => {
  //   //     //     console.log(error);
  //   //     //   });

  //   //     axios
  //   //       .get(url, {
  //   //         headers: {
  //   //           Authorization: `Bearer ${result.value}`,
  //   //         },
  //   //       })
  //   //       .then((response) => {
  //   //         console.log(response.data.value, "data");
  //   //         // message.info(response.data.value.);

  //   //         const z = response.data.value.map((item) => item.Name);
  //   //         message.info(z[0] ?? 0);
  //   //         setD(z);

  //   //         axios
  //   //           .get(messageUrl, {
  //   //             headers: {
  //   //               Authorization: `Bearer ${result.value}`,
  //   //             },
  //   //           })
  //   //           .then((response) => {
  //   //             console.log(response.data.Body.Content);
  //   //           })
  //   //           .catch(() => {
  //   //             message.info(999);
  //   //             console.log(333);
  //   //           });
  //   //         // const data = response.data;
  //   //         // console.log(data, "data");
  //   //         // const list = data.value.map((item) => "data:image/png;base64," + item.ContentBytes);
  //   //         // setList(list);
  //   //       })
  //   //       .catch(() => {
  //   //         message.info(999);
  //   //         console.log(333);
  //   //       });

  //   //     // const response = await axios.get(url, {
  //   //     //   headers: {
  //   //     //     Authorization: `Bearer ${result.value}`,
  //   //     //   },
  //   //     // });
  //   //     // console.log(response.data, "response---");
  //   //   });
  //   // });

  //   const data = Office.context.mailbox.item.conversationId;

  //   console.log(data);

  //   // Office.context.mailbox.item.getAttachmentsAsync((result) => {
  //   //   console.log(result);
  //   // });
  // }, []);

  useEffect(() => {
    const msalConfig = {
      auth: {
        clientId: "bfe5c2cf-8b6a-4694-aa35-b35c0fc08647",
        authority: "https://login.microsoftonline.com/e62ae085-5adb-4fca-9a94-bed260f0f3f3",
      },
    };

    const msalInstance = new PublicClientApplication(msalConfig);

    const scopes = ["user.readwrite"];

    async function getAccessToken() {
      try {
        const response = await msalInstance.loginPopup({
          scopes,
        });

        const tokenResponse = await msalInstance.acquireTokenSilent({
          account: response.account,
          scopes,
        });

        return tokenResponse.accessToken;
      } catch (error) {
        console.log(error);
        throw error;
      }
    }

    getAccessToken();

    // Office.context.mailbox.item.getItemIdAsync((res) => {
    //   console.log(res.value);
    //   Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, async function (result) {
    //     console.log(result.value);
    //     var url = `https://outlook.office.com/api/v2.0/me/messages/${res.value}`;
    //     axios
    //       .get(url, {
    //         headers: {
    //           Authorization: `Bearer ${result.value}`,
    //         },
    //       })
    //       .then((response) => {
    //         console.log(response.data);
    //         // setContent(response.data.BodyPreview);
    //       })
    //       .catch((error) => {
    //         console.log(error);
    //       });
    //   });
    // });
    // Office.context.mailbox.getCallbackTokenAsync(function (result) {
    //   if (result.status === Office.AsyncResultStatus.Succeeded) {
    //     Office.context.mailbox.item.getItemIdAsync((res) => {
    //       if (res.status === Office.AsyncResultStatus.Succeeded) {
    //         const itemId = res.value;
    //         var getRepliesUrl = "https://outlook.office.com/api/v2.0/me/messages/" + itemId + "/replies";
    //         axios
    //           .get(getRepliesUrl, {
    //             headers: {
    //               Authorization: `Bearer ${result.value}`,
    //             },
    //           })
    //           .then((response) => {
    //             console.log(response.data);
    //           })
    //           .catch(() => {
    //             console.log(333);
    //           });
    //       }
    //     });
    //   }
    // });
    // 拆开邮件正文 获取回复
    // Office.context.mailbox.item.body.getAsync("html", (result) => {
    //   var replyBody = result.value;
    //   setContent(replyBody);
    //   var tempDiv = document.createElement("div");
    //   tempDiv.innerHTML = replyBody;
    //   var mailBodyElement = tempDiv.querySelector("div.WordSection1");
    //   if (mailBodyElement) {
    //     // 获取 <div class=WordSection1> 下的所有 <p class=MsoNormal> 元素
    //     var pElements = mailBodyElement.querySelectorAll("p.MsoNormal");
    //     // 遍历每个 <p class=MsoNormal> 元素并获取其内容
    //     for (var i = 0; i < pElements.length; i++) {
    //       var pElement = pElements[i];
    //       var content = pElement.innerHTML;
    //       message.info(content);
    //     }
    //   }
    // });
    // Office.context.mailbox.item.body.getAsync("html", (html) => {
    //   console.log(html.value);
    //   setContent(html.value);
    //   // const tempDiv = document.createElement("div");
    //   // tempDiv.innerHTML = html.value;
    //   // const imgTags = tempDiv.getElementsByTagName("img");
    //   // const promises = [];
    //   // for (let i = 0; i < imgTags.length; i++) {
    //   //   const img = imgTags[i];
    //   //   const src = img.getAttribute("src");
    //   //   var baseUrl = "https://outlook.office.com/";
    //   //   var decodedPath = src.replace(/%7b/g, "{").replace(/%7d/g, "}");
    //   //   promises.push(baseUrl + decodedPath);
    //   // }
    //   //   // setZ(promises);
    //   //   // const tempDiv = document.createElement("div");
    //   //   // tempDiv.innerHTML = html.value;
    //   //   // const imgTags = tempDiv.getElementsByTagName("img");
    //   //   // const promises = [];
    //   //   // for (let i = 0; i < imgTags.length; i++) {
    //   //   //   const img = imgTags[i];
    //   //   //   const src = img.getAttribute("src");
    //   //   //   promises.push(src);
    //   //   // }
    //   //   // setS(promises);
    // });
    // Office.context.mailbox.item.getItemIdAsync((res) => {
    //   console.log(res.value);
    //   // console.log(res.value);
    //   Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, async function (result) {
    //     // var url = `https://outlook.office.com/api/v2.0/me/mailFolders/DeletedItems/messages`;
    //     // axios
    //     //   .get(url, {
    //     //     headers: {
    //     //       Authorization: `Bearer ${result.value}`,
    //     //     },
    //     //   })
    //     //   .then((response) => {
    //     //     console.log(response.data.value);
    //     //   })
    //     //   .catch(() => {
    //     //     message.info(999);
    //     //     console.log(333);
    //     //   });
    //     var url = `https://outlook.office.com/api/v2.0/me/messages/${res.value}/attachments`;
    //     axios
    //       .get(url, {
    //         headers: {
    //           Authorization: `Bearer ${result.value}`,
    //         },
    //       })
    //       .then((response) => {
    //         console.log(response.data.value);
    //         // const z = response.data.value.map((item) => ({
    //         //   // name: item.Name.split(".")[0],
    //         //   name: item.Name,
    //         //   src: item.ContentBytes,
    //         // }));
    //         // setD(z);
    //         // Office.context.mailbox.item.body.getAsync("html", function callback(html) {
    //         //   const tempDiv = document.createElement("div");
    //         //   tempDiv.innerHTML = html.value;
    //         //   const imgTags = tempDiv.getElementsByTagName("img");
    //         //   const promises = [];
    //         //   for (let i = 0; i < imgTags.length; i++) {
    //         //     const img = imgTags[i];
    //         //     const src = img.getAttribute("src");
    //         //     const a = src.split("/").pop().split(".")[0];
    //         //     // console.log(a);
    //         //     promises.push(src);
    //         //     // for (let b = 0; b < z.length; b++) {
    //         //     //   if (i === b) img.setAttribute("src", "data:image/png;base64," + z[b].src);
    //         //     // }
    //         //   }
    //         //   setS(promises);
    //         //   // const updatedHtml = tempDiv.innerHTML;
    //         //   // setContent(updatedHtml);
    //         // });
    //       })
    //       .catch(() => {
    //         message.info(999);
    //         console.log(333);
    //       });
    //   });
    // });
  }, []);

  const translate = () => {
    Office.context.mailbox.item.body.getAsync("html", function callback(result) {
      setContent(JSON.stringify(result.value));
    });
  };

  return { content, setContent, type, translate, src, list, s, d, z };
};
