import { useEffect, useState } from "react";
import { clone } from "ramda";

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
  }, []);

  return { content, setContent };
};
