import { useEffect, useState } from "react";
import { clone } from "ramda";
import { PostTranslate } from "../../api";

export const useAction = () => {
  const [content, setContent] = useState<string>("");

  const [cleanContent, setCleanContent] = useState<string>("");

  const [language, setLanguage] = useState<string>("zh-Tw");

  useEffect(() => {
    // 正文获取
    Office.context.mailbox.item.body.getAsync("html", function callback(result) {
      setContent(result.value);

      const clean = clone(result.value);

      setCleanContent(JSON.stringify(clean));
    });
  }, []);

  const translateContent = () => {
    PostTranslate(cleanContent, language).then((res) => setContent(res));
  };

  const handleChange = (value: string) => {
    setLanguage(value);
  };

  return { content, setContent, translateContent, handleChange, language };
};
