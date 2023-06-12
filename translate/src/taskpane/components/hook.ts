import { useEffect, useState } from "react";
import { clone } from "ramda";
import { PostTranslate } from "../../api";

export const useAction = () => {
  const [content, setContent] = useState<string>("");

  const [cleanContent, setCleanContent] = useState<string>("");

  const [language, setLanguage] = useState<string>("zh-Tw");

  useEffect(() => {
    Office.context.mailbox.item.body.getAsync("html", function callback(result) {
      const clean = clone(result.value);
      setCleanContent(clean);
      replaceImg(clean);
    });
  }, []);

  const replaceImg = (html: string) => {
    const attachments = Office.context.mailbox.item.attachments;
    const tempDiv = document.createElement("div");
    tempDiv.innerHTML = html;

    const imgTags = tempDiv.getElementsByTagName("img");

    const promises = [];

    for (let i = 0; i < imgTags.length; i++) {
      const img = imgTags[i];
      const src = img.getAttribute("src");

      if (src) {
        const attachmentId = attachments[i].id;
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
      })
      .catch((error) => {
        console.error(error);
      });
  };

  const translateContent = async () => {
    const title = await PostTranslate(cleanContent, language).then((res) => {
      console.log(res, "translate");
      return res;
    });

    replaceImg(title);
  };

  const handleChange = (value: string) => {
    setLanguage(value);
  };

  return { content, setContent, translateContent, handleChange, language };
};
