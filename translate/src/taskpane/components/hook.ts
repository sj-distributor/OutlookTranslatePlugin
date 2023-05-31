import { useEffect, useState } from "react";
import { clone } from "ramda";
import { PostTranslate } from "../../api";
import axios from "axios";

export const useAction = () => {
  const [content, setContent] = useState<string>("");

  const [cleanContent, setCleanContent] = useState<string>("");

  const [language, setLanguage] = useState<string>("zh-Tw");

  useEffect(() => {
    // 正文获取
    // Office.context.mailbox.item.body.getAsync("html", {}, function callback(result) {
    //   setContent(result.value);
    //   // console.log(result, "res");

    //   console.log(result.value, "value");

    //   const clean = clone(result.value);

    //   setCleanContent(JSON.stringify(clean));
    // });

    // const item = Office.context.mailbox.item;
    // console.log(item.attachments);
    // let outputString = "";

    // if (item.attachments.length > 0) {
    //   for (let i = 0; i < item.attachments.length; i++) {
    //     const attachment = item.attachments[i];
    //     outputString += "<BR>" + i + ". Name: ";
    //     outputString += attachment.name;
    //     outputString += "<BR>ID: " + attachment.id;
    //     outputString += "<BR>contentType: " + attachment.contentType;
    //     outputString += "<BR>size: " + attachment.size;
    //     outputString += "<BR>attachmentType: " + attachment.attachmentType;
    //     outputString += "<BR>isInline: " + attachment.isInline;
    //   }
    // }

    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
      // console.log(result.value, "111");

      var messageId = Office.context.mailbox.item.itemId;

      const attachmentId = Office.context.mailbox.item.attachments[0].id;

      console.log(messageId, attachmentId, "attachmentId");

      var url = "https://outlook.office.com/api/v1.0/me/messages" + messageId + "/attachments/" + attachmentId;

      axios
        .get(url, {
          headers: {
            Authorization: "Bearer " + result.value,
            Accept: "application/json",
          },
        })
        .then((response) => {
          return response.data as string;
        })
        .catch((err) => console.log(err));

      // fetch({
      //   url: url,
      //   type: "GET",
      //   headers: {
      //     Authorization: "Bearer " + result.value,
      //     Accept: "application/json",
      //   },
      // }).then((res) => console.log(res));
    });
  }, []);

  const translateContent = () => {
    PostTranslate(cleanContent, language).then((res) => {
      console.log(res, "translate");
      setContent(res);
    });
  };

  const handleChange = (value: string) => {
    setLanguage(value);
  };

  return { content, setContent, translateContent, handleChange, language };
};
