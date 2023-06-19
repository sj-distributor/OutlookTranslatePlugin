import { useEffect, useState } from "react";
import axios from "axios";
import { message } from "antd";
import { PostTranslate } from "../../api";

export const useAction = () => {
  const [content, setContent] = useState<string>("");

  // // 辅助函数：生成随机的 code verifier
  // function generateCodeVerifier() {
  //   // 生成随机的字节数组
  //   const array = new Uint8Array(32);
  //   window.crypto.getRandomValues(array);

  //   // 将字节数组转换为 base64 字符串
  //   let codeVerifier = "";
  //   for (let i = 0; i < array.length; i++) {
  //     codeVerifier += String.fromCharCode(array[i]);
  //   }
  //   codeVerifier = btoa(codeVerifier);

  //   return codeVerifier;
  // }

  // // 辅助函数：根据 code verifier 生成 code challenge
  // function generateCodeChallenge(codeVerifier) {
  //   // 使用 SHA-256 哈希算法计算 code verifier 的哈希值
  //   const hash = SHA256(codeVerifier);

  //   // 将哈希值转换为 base64 URL 编码的字符串
  //   let codeChallenge = hash.toString(enc.Base64);
  //   codeChallenge = codeChallenge.replace(/=/g, "").replace(/\+/g, "-").replace(/\//g, "_");

  //   return codeChallenge;
  // }

  // useEffect(() => {
  //   // const codeVerifier = generateCodeVerifier();
  //   // // 使用 code verifier 生成 code challenge
  //   // const codeChallenge = generateCodeChallenge(codeVerifier);
  //   // const clientId = "38883baa-ec1c-4e93-b4db-bebbe79b5807";
  //   // const redirectUri = "http://localhost:3000/edit.html";
  //   // const scopes = ["User.Read", "Mail.ReadWrite"];
  //   // const scopesString = scopes.join(" ");
  //   // const authorizationUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${clientId}&response_type=code&redirect_uri=${encodeURIComponent(
  //   //   redirectUri
  //   // )}&scope=${encodeURIComponent(scopesString)}&code_challenge=${encodeURIComponent(
  //   //   codeChallenge
  //   // )}&code_challenge_method=S256`;
  //   // const handleMessage = (event) => {
  //   //   // 确保消息是从打开的弹出窗口发送的
  //   //   console.log(event);
  //   //   if (event.source === popupRef.current) {
  //   //     // 检查消息中是否包含授权码
  //   //     if (event.data && event.data.code) {
  //   //       const authorizationCode = event.data.code;
  //   //       console.log(authorizationCode);
  //   //       // 在这里处理授权码，可以将其传递给主窗口进行后续处理
  //   //       window.opener.postMessage({ code: authorizationCode }, "*");
  //   //       // 关闭弹出窗口
  //   //       popupRef.current.close();
  //   //     } else {
  //   //       // 处理错误情况
  //   //       console.error("无法获取授权码");
  //   //     }
  //   //   }
  //   // };
  //   // window.addEventListener("message", handleMessage);
  //   // return () => {
  //   //   window.removeEventListener("message", handleMessage);
  //   // };
  // }, []);

  // useEffect(() => {
  //   Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, async (result) => {
  //     console.log(result.value);
  //   });
  //   try {
  //     const fun = async () => {
  //       console.log(123);
  //       const accessToken = await OfficeRuntime.auth.getAccessToken({
  //         allowSignInPrompt: false,
  //         allowConsentPrompt: false,
  //         forMSGraphAccess: true,
  //       });

  //       console.log(accessToken, "-----");
  //     };

  //     fun();
  //   } catch (error) {
  //     console.log("Error obtaining token", error);
  //   }
  // }, []);

  // useEffect(() => {
  //   let dialog; // Declare dialog as global for use in later functions.
  //   const codeVerifier = generateCodeVerifier();
  //   // 使用 code verifier 生成 code challenge
  //   const codeChallenge = generateCodeChallenge(codeVerifier);
  //   const clientId = "38883baa-ec1c-4e93-b4db-bebbe79b5807";
  //   const redirectUri = "https://localhost:3000/";
  //   const scopes = ["User.Read", "Mail.ReadWrite"];
  //   const scopesString = scopes.join(" ");
  //   const authorizationUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${clientId}&response_type=code&redirect_uri=${encodeURIComponent(
  //     redirectUri
  //   )}&scope=${encodeURIComponent(scopesString)}&code_challenge=${encodeURIComponent(
  //     codeChallenge
  //   )}&code_challenge_method=S256`;

  //   Office.context.ui.displayDialogAsync(authorizationUrl, { height: 30, width: 20 }, function (asyncResult) {
  //     console.log(asyncResult?.value);
  //     message.info(asyncResult);
  //   });

  //   // const profileMessage = {
  //   //   name: "1",
  //   //   email: "1",
  //   // };
  //   // Office.context.ui.messageParent(JSON.stringify(profileMessage));
  // }, []);

  // useEffect(() => {
  //   // 根据您的应用程序配置设置相应的参数值
  //   const clientId = "YOUR_CLIENT_ID";
  //   const graphScopes = ["User.Read", "Mail.Read"]; // 需要的权限范围

  //   // 初始化 MSAL.js PublicClientApplication
  //   const msalConfig = {
  //     auth: {
  //       clientId: clientId,
  //       authority: "https://login.microsoftonline.com/common",
  //       redirectUri: window.location.origin,
  //     },
  //   };
  //   const pca = new PublicClientApplication(msalConfig);

  //   // 注册 Teams SSO 回调函数
  //   microsoftTeams.initialize();
  //   microsoftTeams.authentication.registerForAuthStatus(async (authStatus) => {
  //     if (authStatus === microsoftTeams.AuthenticationStatus.Authenticated) {
  //       try {
  //         // 使用 MSAL.js 获取 Microsoft Graph 令牌
  //         const accounts = pca.getAllAccounts();
  //         const tokenRequest = {
  //           scopes: graphScopes,
  //           account: accounts[0],
  //         };
  //         const response = await pca.acquireTokenSilent(tokenRequest);

  //         // 在这里处理获取到的 Microsoft Graph 令牌
  //         const accessToken = response.accessToken;
  //         // ...
  //         // 执行您的后续操作
  //       } catch (error) {
  //         console.log(error);
  //         // 处理获取令牌失败的情况
  //       }
  //     }
  //   });

  //   // 启动 Teams SSO 身份验证流程
  //   microsoftTeams.authentication.authenticate({
  //     url: window.location.origin + "/auth-endpoint", // 后端服务的身份验证端点
  //     successCallback: () => {
  //       // 验证成功回调
  //     },
  //     failureCallback: (reason) => {
  //       console.log(reason);
  //       // 处理验证失败回调
  //     },
  //   });
  // }, []);

  // useEffect(() => {
  //   Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, async (result) => {
  //     if (result.status === Office.AsyncResultStatus.Succeeded) {
  //       var accessToken = result.value;
  //       var graphEndpoint = "/api/common";

  //       fetch(graphEndpoint, {
  //         headers: {
  //           Authorization: `Bearer ${accessToken}`,
  //         },
  //       })
  //         .then((response) => response.json())
  //         .then((userData) => {
  //           // 在这里处理从 Microsoft Graph 返回的用户数据
  //           console.log(userData);
  //         })
  //         .catch((error) => {
  //           console.log(error);
  //         });
  //     } else {
  //       console.log("Failed to get access token:", result.error);
  //     }
  //   });
  // }, []);

  // useEffect(() => {
  //   axios
  //     .get(
  //       `/api/common/oauth2/v2.0/authorize?client_id=38883baa-ec1c-4e93-b4db-bebbe79b5807&response_type=code&redirect_uri=https://localhost:3000/edit.html&response_mode=query&scope=user.read&state=12345`
  //     )
  //     .then((res) => console.log(res))
  //     .catch((err) => console.log(err));
  // }, []);

  // useEffect(() => {
  //   const codeVerifier = generateCodeVerifier();

  //   //使用 code verifier 生成 code challenge
  //   const codeChallenge = generateCodeChallenge(codeVerifier);
  //   const clientId = "38883baa-ec1c-4e93-b4db-bebbe79b5807";
  //   const redirectUri = "https://localhost:3000/";
  //   const scopes = ["User.Read", "Mail.ReadWrite"];
  //   const scopesString = scopes.join(" ");

  //   const authorizationUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${clientId}&response_type=code&redirect_uri=${encodeURIComponent(
  //     redirectUri
  //   )}&scope=${encodeURIComponent(scopesString)}&code_challenge=${encodeURIComponent(
  //     codeChallenge
  //   )}&code_challenge_method=S256`;

  //   const popup = window.open(authorizationUrl, "_blank");

  //   window.addEventListener("message", (event) => {
  //     // 确保消息是从打开的弹出窗口发送的
  //     if (event.source === popup) {
  //       // 检查消息中是否包含授权码
  //       if (event.data && event.data.code) {
  //         const authorizationCode = event.data.code;
  //         console.log(authorizationCode);
  //         // 在这里处理授权码，可以将其传递给主窗口进行后续处理
  //         window.opener.postMessage({ code: authorizationCode }, "*");
  //         // 关闭弹出窗口
  //         popup.close();
  //       } else {
  //         // 处理错误情况
  //         console.error("无法获取授权码");
  //       }
  //     }
  //   });

  //   // window.open(authorizationUrl, "_blank");

  //   // axios
  //   //   .get(authorizationUrl)
  //   //   .then((response) => {
  //   //     // 提取返回 URL 中的参数
  //   //     const urlParams = queryString.parse(response.request.responseURL);
  //   //     const authorizationCode = urlParams.code;

  //   //     // 在这里可以使用授权码进行后续操作
  //   //     console.log(authorizationCode);
  //   //   })
  //   //   .catch((error) => {
  //   //     console.error("获取授权码失败:", error);
  //   //   });

  //   // const clientId = "38883baa-ec1c-4e93-b4db-bebbe79b5807";
  //   // const clientSecret = "xHk8Q~5SHS1uKfYFoA31i0p3Y2m.leHFPgWAmcy.";
  //   // const redirectUri = "https://localhost:3000/edit.html";
  //   // const authorizationCode = "AUTHORIZATION_CODE";
  //   // const scopes = ["User.Read", "Mail.ReadWrite"];

  //   // const tokenUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
  //   // const data = {
  //   //   client_id: clientId,
  //   //   client_secret: clientSecret,
  //   //   redirect_uri: redirectUri,
  //   //   code: authorizationCode,
  //   //   grant_type: "authorization_code",
  //   //   scope: scopes,
  //   // };

  //   // axios
  //   //   .post(tokenUrl, data)
  //   //   .then((response) => {
  //   //     console.log(response, "response");
  //   //     // 使用访问令牌进行其他操作
  //   //     // ...
  //   //   })
  //   //   .catch((error) => {
  //   //     console.error("获取访问令牌失败:", error);
  //   //   });
  // }, []);

  // const [type, setType] = useState<string>("html");

  // const [list, setList] = useState<string[]>([]);

  // const [src, setSrc] = useState<string>("");

  // const [s, setS] = useState<string[]>([]);

  // const [d, setD] = useState<
  //   {
  //     name: string;
  //     src: string;
  //   }[]
  // >([]);

  // const [z, setZ] = useState<string[]>([]);

  // const [cleanContent, setCleanContent] = useState<string>("");

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

  //   Office.context.mailbox.item.body.getAsync("text", (result) => {
  //     console.log(result.value);
  //   });
  // };

  // useEffect(() => {
  //   // Office.context.mailbox.item.getItemIdAsync((res) => {
  //   //   console.log(res.value);
  //   //   Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, async function (result) {
  //   //     console.log(result.value);
  //   //     var url = `https://outlook.office.com/api/v2.0/me/messages/${res.value}`;
  //   //     axios
  //   //       .get(url, {
  //   //         headers: {
  //   //           Authorization: `Bearer ${result.value}`,
  //   //         },
  //   //       })
  //   //       .then((response) => {
  //   //         console.log(response.data);
  //   //         // setContent(response.data.BodyPreview);
  //   //       })
  //   //       .catch((error) => {
  //   //         console.log(error);
  //   //       });
  //   //   });
  //   // });
  //   // 拆开邮件正文 获取回复
  //   // Office.context.mailbox.item.body.getAsync("html", (result) => {
  //   //   var replyBody = result.value;
  //   //   setContent(replyBody);
  //   //   var tempDiv = document.createElement("div");
  //   //   tempDiv.innerHTML = replyBody;
  //   //   var mailBodyElement = tempDiv.querySelector("div.WordSection1");
  //   //   if (mailBodyElement) {
  //   //     // 获取 <div class=WordSection1> 下的所有 <p class=MsoNormal> 元素
  //   //     var pElements = mailBodyElement.querySelectorAll("p.MsoNormal");
  //   //     // 遍历每个 <p class=MsoNormal> 元素并获取其内容
  //   //     for (var i = 0; i < pElements.length; i++) {
  //   //       var pElement = pElements[i];
  //   //       var content = pElement.innerHTML;
  //   //       message.info(content);
  //   //     }
  //   //   }
  //   // });

  //   Office.context.mailbox.item.getItemIdAsync((res) => {
  //     console.log(res.value);
  //     // console.log(res.value);
  //     Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, async function (result) {
  //       var url = `https://outlook.office.com/api/v2.0/me/messages/${res.value}`;
  //       axios
  //         .get(url, {
  //           headers: {
  //             Authorization: `Bearer ${result.value}`,
  //           },
  //         })
  //         .then((response) => {
  //           console.log(response.data.Body.Content);
  //           // const z = response.data.value.map((item) => ({
  //           //   // name: item.Name.split(".")[0],
  //           //   name: item.Name,
  //           //   src: item.ContentBytes,
  //           // }));
  //           // setD(z);
  //           // Office.context.mailbox.item.body.getAsync("html", function callback(html) {
  //           //   const tempDiv = document.createElement("div");
  //           //   tempDiv.innerHTML = html.value;
  //           //   const imgTags = tempDiv.getElementsByTagName("img");
  //           //   const promises = [];
  //           //   for (let i = 0; i < imgTags.length; i++) {
  //           //     const img = imgTags[i];
  //           //     const src = img.getAttribute("src");
  //           //     const a = src.split("/").pop().split(".")[0];
  //           //     // console.log(a);
  //           //     promises.push(src);
  //           //     // for (let b = 0; b < z.length; b++) {
  //           //     //   if (i === b) img.setAttribute("src", "data:image/png;base64," + z[b].src);
  //           //     // }
  //           //   }
  //           //   setS(promises);
  //           //   // const updatedHtml = tempDiv.innerHTML;
  //           //   // setContent(updatedHtml);
  //           // });
  //         })
  //         .catch(() => {
  //           message.info(999);
  //           console.log(333);
  //         });
  //     });
  //   });

  //   convert();
  //   // Office.context.mailbox.item.body.getAsync("html", (html) => {
  //   //   console.log(html.value);
  //   //   setContent(html.value);
  //   //   // const tempDiv = document.createElement("div");
  //   //   // tempDiv.innerHTML = html.value;
  //   //   // const imgTags = tempDiv.getElementsByTagName("img");
  //   //   // const promises = [];
  //   //   // for (let i = 0; i < imgTags.length; i++) {
  //   //   //   const img = imgTags[i];
  //   //   //   const src = img.getAttribute("src");
  //   //   //   var baseUrl = "https://outlook.office.com/";
  //   //   //   var decodedPath = src.replace(/%7b/g, "{").replace(/%7d/g, "}");
  //   //   //   promises.push(baseUrl + decodedPath);
  //   //   // }
  //   //   //   // setZ(promises);
  //   //   //   // const tempDiv = document.createElement("div");
  //   //   //   // tempDiv.innerHTML = html.value;
  //   //   //   // const imgTags = tempDiv.getElementsByTagName("img");
  //   //   //   // const promises = [];
  //   //   //   // for (let i = 0; i < imgTags.length; i++) {
  //   //   //   //   const img = imgTags[i];
  //   //   //   //   const src = img.getAttribute("src");
  //   //   //   //   promises.push(src);
  //   //   //   // }
  //   //   //   // setS(promises);
  //   // });
  //   // Office.context.mailbox.item.getItemIdAsync((res) => {
  //   //   console.log(res.value);
  //   //   // console.log(res.value);
  //   //   Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, async function (result) {
  //   //     // var url = `https://outlook.office.com/api/v2.0/me/mailFolders/DeletedItems/messages`;
  //   //     // axios
  //   //     //   .get(url, {
  //   //     //     headers: {
  //   //     //       Authorization: `Bearer ${result.value}`,
  //   //     //     },
  //   //     //   })
  //   //     //   .then((response) => {
  //   //     //     console.log(response.data.value);
  //   //     //   })
  //   //     //   .catch(() => {
  //   //     //     message.info(999);
  //   //     //     console.log(333);
  //   //     //   });
  //   //     var url = `https://outlook.office.com/api/v2.0/me/messages/${res.value}/attachments`;
  //   //     axios
  //   //       .get(url, {
  //   //         headers: {
  //   //           Authorization: `Bearer ${result.value}`,
  //   //         },
  //   //       })
  //   //       .then((response) => {
  //   //         console.log(response.data.value);
  //   //         // const z = response.data.value.map((item) => ({
  //   //         //   // name: item.Name.split(".")[0],
  //   //         //   name: item.Name,
  //   //         //   src: item.ContentBytes,
  //   //         // }));
  //   //         // setD(z);
  //   //         // Office.context.mailbox.item.body.getAsync("html", function callback(html) {
  //   //         //   const tempDiv = document.createElement("div");
  //   //         //   tempDiv.innerHTML = html.value;
  //   //         //   const imgTags = tempDiv.getElementsByTagName("img");
  //   //         //   const promises = [];
  //   //         //   for (let i = 0; i < imgTags.length; i++) {
  //   //         //     const img = imgTags[i];
  //   //         //     const src = img.getAttribute("src");
  //   //         //     const a = src.split("/").pop().split(".")[0];
  //   //         //     // console.log(a);
  //   //         //     promises.push(src);
  //   //         //     // for (let b = 0; b < z.length; b++) {
  //   //         //     //   if (i === b) img.setAttribute("src", "data:image/png;base64," + z[b].src);
  //   //         //     // }
  //   //         //   }
  //   //         //   setS(promises);
  //   //         //   // const updatedHtml = tempDiv.innerHTML;
  //   //         //   // setContent(updatedHtml);
  //   //         // });
  //   //       })
  //   //       .catch(() => {
  //   //         message.info(999);
  //   //         console.log(333);
  //   //       });
  //   //   });
  //   // });
  // }, []);

  // useEffect(() => {
  //   !!content &&
  //     PostTranslate(content, language).then((res) => {
  //       // pElement.innerHTML = res;
  //       // message.info(JSON.parse(res));
  //       message.info(res);

  //       setContent(res);
  //     });
  // }, [content]);

  // const [accessToken, setAccessToken] = useState("");

  // const [a, setA] = useState<string>("");

  // useEffect(() => {
  //   const msalConfig = {
  //     auth: {
  //       clientId: "38883baa-ec1c-4e93-b4db-bebbe79b5807",
  //       authority: "https://login.microsoftonline.com/e62ae085-5adb-4fca-9a94-bed260f0f3f3",
  //     },
  //   };
  //   const msalInstance = new PublicClientApplication(msalConfig);
  //   // 请求的权限范围
  //   const request = {
  //     scopes: ["https://graph.microsoft.com/.default"],
  //   };
  //   message.info(window.location.href);
  //   setA(window.location.href);
  //   msalInstance
  //     .loginPopup(request)
  //     .then((response) => {
  //       // 从响应中提取访问令牌
  //       const accessToken = response.accessToken;
  //       console.log(accessToken, "accessToken---");
  //       setAccessToken(accessToken);
  //     })
  //     .catch((error) => {
  //       console.log(error);
  //       message.error(error.message);
  //     });
  // }, []);

  // const convert = () => {
  //   // Office.context.mailbox.item.body.getAsync("html", async (result) => {
  //   //   var replyBody = result.value;
  //   //   // setContent(replyBody);
  //   //   var tempDiv = document.createElement("div");
  //   //   tempDiv.innerHTML = replyBody;
  //   //   var mailBodyElement = tempDiv.querySelector("div.WordSection1");
  //   //   if (mailBodyElement) {
  //   //     // 获取 <div class=WordSection1> 下的所有 <p class=MsoNormal> 元素
  //   //     var pElements = mailBodyElement.querySelectorAll("p.MsoNormal");
  //   //     // 遍历每个 <p class=MsoNormal> 元素并获取其内容
  //   //     for (var i = 0; i < pElements.length; i++) {
  //   //       var pElement = pElements[i];
  //   //       var content = pElement.innerHTML;
  //   //       if (content.indexOf('<a name="_MailOriginal">') !== -1) {
  //   //         break;
  //   //       }
  //   //       await PostTranslate(content, language).then((res) => {
  //   //         pElement.innerHTML = res;
  //   //         // message.info(JSON.parse(res));
  //   //       });
  //   //       setContent(tempDiv.innerHTML);
  //   //     }
  //   //   }
  //   // });
  // };

  // useEffect(() => {
  //   // Office.context.mailbox.item.getItemIdAsync((res) => {
  //   //   Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, async function (result) {
  //   //     var url = `https://outlook.office.com/api/v2.0/me/messages/${res.value}`;
  //   //     axios
  //   //       .get(url, {
  //   //         headers: {
  //   //           Authorization: `Bearer ${result.value}`,
  //   //         },
  //   //       })
  //   //       .then((response) => {
  //   //         console.log(response.data.Body.Content);
  //   //       })
  //   //       .catch(() => {
  //   //         message.info(999);
  //   //         console.log(333);
  //   //       });
  //   //   });
  //   // });

  //   Office.context.mailbox.item.body.getAsync("html", async (result) => {
  //     setCleanContent(result.value);
  //     setContent(result.value);
  //     // convert(result.value);
  //   });

  //   // convert();

  //   // Office.context.mailbox.item.body.getAsync("html", async (result) => {
  //   //   console.log(result.value);
  //   //   setContent(result.value);

  //   //   // await PostTranslate(result.value, language).then((res) => {
  //   //   //   // pElement.innerHTML = res;
  //   //   //   message.info(JSON.parse(res));
  //   //   //   console.log(res);
  //   //   //   setContent(res);
  //   //   // });
  //   // });
  // }, []);

  const [language, setLanguage] = useState<string>("zh-Tw");

  useEffect(() => {
    if (localStorage.getItem("language")) {
      const cachedData = localStorage.getItem("language");
      setLanguage(cachedData);
    }
  }, []);

  // useEffect(() => {
  //   // 定义应用程序的客户端ID、租户ID和重定向URL
  //   const clientId = "01010e21-8a9a-4221-9044-e1c4c91512e7";
  //   const tenantId = "e62ae085-5adb-4fca-9a94-bed260f0f3f3";
  //   const clientSecret = "~f08Q~G~5iOa1dWJOhmUWnyAFZU0CM_yTGQm3a1_";
  //   const redirectUri = "https://localhost:3000/";

  //   // 定义要请求的权限范围
  //   const scopes = ["https://graph.microsoft.com/.default"];

  //   const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${clientId}&response_type=code&redirect_uri=${encodeURIComponent(
  //     redirectUri
  //   )}&scope=${encodeURIComponent(scopes.join(" "))}&code_challenge=${encodeURIComponent(
  //     generateCodeChallenge(generateCodeVerifier())
  //   )}&code_challenge_method=S256`;

  //   // window.location.href = authUrl;

  //   const loginWindow = window.open(authUrl, "Login", "width=600,height=400");

  //   // 监听窗口加载完成事件
  //   loginWindow.addEventListener("load", function () {
  //     var loginWindowUrl = loginWindow.location.href;
  //     console.log("Login window URL:", loginWindowUrl);
  //   });
  // }, []);

  // 获取token
  // useEffect(() => {
  //   const clientId = "01010e21-8a9a-4221-9044-e1c4c91512e7";
  //   const redirectUri = "https://localhost:3000";
  //   const scopes = ["user.read", "mail.read"];

  //   const msalConfig = {
  //     auth: {
  //       clientId: clientId,
  //       redirectUri: redirectUri,
  //       responseType: "code",
  //     },
  //   };

  //   const msalInstance = new PublicClientApplication(msalConfig);
  //   msalInstance
  //     .loginPopup({ scopes: scopes })
  //     .then((response) => {
  //       // 登录成功，获取访问令牌
  //       const accessToken = response.accessToken;
  //       console.log("Access token:", accessToken);

  //       message.info("success");

  //       // 使用访问令牌进行后续操作，如调用 Microsoft Graph API
  //     })
  //     .catch((error) => {
  //       // 处理登录错误
  //       console.error("Login error:", error);
  //       message.info("999");
  //     });
  // }, []);

  // useEffect(() => {
  //   (async () => {
  //     try {
  //       const accessToken = await Office.auth.getAccessToken({
  //         allowSignInPrompt: true,
  //         allowConsentPrompt: true,
  //         forMSGraphAccess: true,
  //       });
  //       message.info(accessToken);
  //     } catch (error) {
  //       message.error(error.message);
  //     }
  //   })();
  // }, []);

  // useEffect(() => {
  //   const clientId = "01010e21-8a9a-4221-9044-e1c4c91512e7";
  //   const redirectUri = "https://localhost:3000";
  //   const scopes = ["user.read", "mail.read"];

  //   // 构建授权 URL
  //   const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${clientId}&response_type=token&redirect_uri=${encodeURIComponent(
  //     redirectUri
  //   )}&scope=${encodeURIComponent(scopes.join(" "))}`;

  //   Office.context.ui.displayDialogAsync(authUrl, { height: 50, width: 50 }, dialogCallback);
  // }, []);

  // useEffect(() => {
  //   const clientId = "01010e21-8a9a-4221-9044-e1c4c91512e7";
  //   const redirectUri = "https://localhost:3000";
  //   const scopes = ["user.read", "mail.read"];

  //   // 构建授权 URL
  //   const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${clientId}&response_type=token&redirect_uri=${encodeURIComponent(
  //     redirectUri
  //   )}&scope=${encodeURIComponent(scopes.join(" "))}`;

  //   Office.context.ui.displayDialogAsync(authUrl, { height: 50, width: 50 }, function (result) {
  //     const dialog = result.value;

  //     dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (args) {
  //       if ("error" in args) {
  //         console.error("Dialog Error:", args.error);
  //         // 处理对话框错误
  //       } else if (args.message && args.origin) {
  //         // 处理从对话框传递回来的消息和来源

  //         const message = JSON.parse(args.message);
  //         if (message && message.code) {
  //           const authCode = message.code;
  //           // getToken(authCode);
  //           message.info(authCode);
  //         }
  //       }

  //       dialog.close();
  //     });
  //   });
  // }, []);

  // useEffect(() => {
  //   const clientId = "01010e21-8a9a-4221-9044-e1c4c91512e7";
  //   const clientSecret = "~f08Q~G~5iOa1dWJOhmUWnyAFZU0CM_yTGQm3a1_";
  //   const scope = "https://graph.microsoft.com/.default";

  //   const getToken = async () => {
  //     try {
  //       const response = await axios.post("https://login.microsoftonline.com/common/oauth2/v2.0/token", {
  //         client_id: clientId,
  //         scope: scope,
  //         client_secret: clientSecret,
  //         grant_type: "client_credentials",
  //       });

  //       const accessToken = response.data.access_token;
  //       console.log("Access Token:", accessToken);
  //     } catch (error) {
  //       console.error("Error obtaining token:", error);
  //     }
  //   };

  //   getToken();
  // }, []);

  const isHTML = (str: string) => {
    const doc = new DOMParser().parseFromString(str, "text/html");
    return Array.from(doc.body.childNodes).some((node) => node.nodeType === 1);
  };

  const convert = async () => {
    Office.context.mailbox.item.body.getAsync("html", async (result) => {
      var replyBody = result.value;
      var tempDiv = document.createElement("div");
      tempDiv.innerHTML = replyBody;
      var mailBodyElement = tempDiv.querySelector("div.WordSection1");
      if (mailBodyElement) {
        // 获取 <div class=WordSection1> 下的所有 <p class=MsoNormal> 元素
        var pElements = mailBodyElement.querySelectorAll("p.MsoNormal");
        // 遍历每个 <p class=MsoNormal> 元素并获取其内容
        for (let i = 0; i < pElements.length; i++) {
          const content = pElements[i].innerHTML;
          if (content.indexOf('<a name="_MailOriginal">') !== -1) {
            break;
          }
          try {
            const change = content.replace(/"/g, "'");
            const translatedContent = await PostTranslate(change, language, isHTML(change));
            pElements[i].innerHTML = translatedContent;
          } catch (error) {
            message.error(1);
          }
        }
        setContent(tempDiv.innerHTML);
        Office.context.mailbox.item.body.setAsync(tempDiv.innerHTML, { coercionType: "html" }, async () => {
          Office.context.mailbox.item.body.getAsync("html", (result) => {
            setContent(result.value);
          });
        });
      }
    });
  };

  const translate = () => {
    convert();
  };

  const handleChange = (value: string) => {
    setLanguage(value);
    localStorage.setItem("language", value);
  };

  return { content, setContent, translate, language, handleChange };
};
