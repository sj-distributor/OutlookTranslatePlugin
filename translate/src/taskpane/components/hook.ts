import { useEffect, useState } from "react";
import { clone } from "ramda";
import { PostTranslate } from "../../api";

export const useAction = () => {
  const [content, setContent] = useState<string>("");

  const [cleanContent, setCleanContent] = useState<string>("");

  useEffect(() => {
    // 正文获取
    Office.context.mailbox.item.body.getAsync("html", function callback(result) {
      console.log(result, "html");
      setContent(result.value);

      const clean = clone(result.value);

      setCleanContent(clean);
    });

    Office.context.mailbox.item.body.getAsync("text", function callback(result) {
      console.log(result, "text");
      // setContent(result.value);

      // const clean = clone(result.value);

      // setCleanContent(clean);
    });

    // console.log(Office.context.mailbox.item.subject);
    // // 标题获取
    // setB(Office.context.mailbox.item.subject);

    // getTranslate();

    // Office.context.mailbox.item.
    // PostTranslate().then((res) => console.log(res, "res"));
  }, []);

  const translateContent = () => {
    PostTranslate(cleanContent).then((res) => console.log(res, "res"));
  };

  const handleChange = (value: string) => {
    console.log(`selected ${value}`);
  };

  return { content, setContent, translateContent, handleChange };
};
