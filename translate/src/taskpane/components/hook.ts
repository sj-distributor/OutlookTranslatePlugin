import { useEffect, useState } from "react";
import { PostTranslate } from "../../api";
import { message } from "antd";
import axios from "axios";

export const useAction = () => {
  const [content, setContent] = useState<string>("");

  useEffect(() => {
    replaceImg();
  }, []);

  // 根据替换图片src路径
  const replaceImg = () => {
    let mailBodyHtml = "";

    const attachments = [];

    const id = Office.context.mailbox.item.itemId;

    try {
      Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, async (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const token = result.value;
          var apiHtmlUrl = Office.context.mailbox.restUrl + "/v2.0/me/messages/" + id;

          await axios
            .get(apiHtmlUrl, {
              headers: {
                Authorization: "Bearer " + token,
              },
            })
            .then(async (response) => {
              const draftData = response.data;

              // 去除字符串的双引号
              const newString = draftData.Body.Content.replace(/"/g, "'");

              const res = await PostTranslate(newString, "zh-TW");

              if (res.msg) {
                mailBodyHtml = draftData.Body.Content;
              } else {
                mailBodyHtml = res;
              }
            });

          const tempDiv = document.createElement("div");

          tempDiv.innerHTML = mailBodyHtml;

          const imgTags = tempDiv.getElementsByTagName("img");

          if (imgTags.length <= 0) {
            setContent(mailBodyHtml);
          } else {
            const apiAttachmentsUrl = Office.context.mailbox.restUrl + "/v2.0/me/messages/" + id + "/attachments";

            await axios
              .get(apiAttachmentsUrl, {
                headers: {
                  Authorization: "Bearer " + token,
                },
              })
              .then((response) => {
                // 获取附件list
                response.data.value.map((item) => {
                  attachments.push({
                    id: item.Id,
                    contentId: item.ContentId,
                    base64: item.ContentBytes,
                  });
                });

                for (let i = 0; i < imgTags.length; i++) {
                  const img = imgTags[i];
                  const src = img.getAttribute("src");
                  const split = src.split("cid:")[1];

                  const index = attachments.findIndex((item) => item.contentId === split);

                  index > -1 && img.setAttribute("src", "data:image/png;base64," + attachments[index].base64);
                }
                const updatedHtml = tempDiv.innerHTML;
                setContent(updatedHtml);
                if (tempDiv.parentNode) {
                  tempDiv.parentNode.removeChild(tempDiv);
                }
              });
          }
        }
      });
    } catch (error) {
      message.error(error.message);
    }
  };

  return { content };
};
