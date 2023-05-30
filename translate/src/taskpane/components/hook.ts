import { useEffect, useState } from "react";
import { clone } from "ramda";
import { PostTranslate } from "../../api";

export const useAction = () => {
  const [content, setContent] = useState<string>("");

  const [cleanContent, setCleanContent] = useState<string>("");

  useEffect(() => {
    // 正文获取
    Office.context.mailbox.item.body.getAsync("html", function callback(result) {
      setContent(result.value);

      const clean = clone(result.value);

      setCleanContent(clean);
    });

    // console.log(Office.context.mailbox.item.subject);
    // // 标题获取
    // setB(Office.context.mailbox.item.subject);

    // getTranslate();

    // Office.context.mailbox.item.
    // PostTranslate().then((res) => console.log(res, "res"));
  }, []);

  const translateContent = () => {
    const text = JSON.stringify(cleanContent);
    console.log(text, "text-translate");
    PostTranslate(text).then((res) => console.log(res, "res"));
  };

  return { content, setContent, translateContent };
};
