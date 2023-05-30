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
      console.log(JSON.stringify(result.value), "html");

      setContent(result.value);

      const clean = clone(result.value);

      setCleanContent(JSON.stringify(clean));
    });
  }, []);

  const translateContent = () => {
    PostTranslate(cleanContent).then((res) => console.log(res, "res"));
  };

  const handleChange = (value: string) => {
    console.log(`selected ${value}`);
  };

  return { content, setContent, translateContent, handleChange };
};
