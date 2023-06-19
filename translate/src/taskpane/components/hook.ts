import { useEffect, useState } from "react";
import { PostTranslate } from "../../api";

export const useAction = () => {
  const [content, setContent] = useState<string>("");

  useEffect(() => {
    Office.context.mailbox.item.body.getAsync("html", async function callback(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const res = await PostTranslate(result.value, "zh-TW");
        if (res.msg) {
          replaceImg(result.value);
        } else {
          replaceImg(res);
        }
      }
    });
  }, []);

  // 根据替换图片src路径
  const replaceImg = (html: string) => {
    const attachments = Office.context.mailbox.item.attachments;
    const tempDiv = document.createElement("div");
    tempDiv.innerHTML = html;

    const imgTags = tempDiv.getElementsByTagName("img");

    const promises = [];

    if (attachments.length <= 0 || imgTags.length <= 0) {
      setContent(html);
    } else {
      for (let i = 0; i < imgTags.length; i++) {
        const img = imgTags[i];
        const src = img.getAttribute("src");
        if (src) {
          const attachmentId = attachments[i]?.id;
          const promise = new Promise((resolve, reject) => {
            Office.context.mailbox.item.getAttachmentContentAsync(attachmentId, (result) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value.content);
              } else {
                reject(new Error("Failed to get attachment content."));
              }
            });
          });

          promises.push(promise);
        }
      }

      Promise.all(promises)
        .then((base64Contents) => {
          for (let i = 0; i < imgTags.length; i++) {
            const img = imgTags[i];
            img.setAttribute("src", "data:image/png;base64," + base64Contents[i]);
          }

          const updatedHtml = tempDiv.innerHTML;
          setContent(updatedHtml);
          if (tempDiv.parentNode) {
            tempDiv.parentNode.removeChild(tempDiv);
          }
        })
        .catch((error) => {
          console.error(error);
        });
    }
  };

  return { content };
};
