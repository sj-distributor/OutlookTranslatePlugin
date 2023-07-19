import { useEffect, useState } from "react";
import { message } from "antd";
import axios from "axios";
import { useDebounce } from "ahooks";
import * as wangEditor from "@wangeditor/editor";
import { PostTranslate } from "../../api";

enum ApiType {
  Office,
  Rest,
}

enum Language {
  Chinese,
  English,
  Spanish,
}

const LanguageType = {
  [Language.Chinese]: "zh-TW",
  [Language.English]: "en",
  [Language.Spanish]: "es",
};

const baseHtml =
  "<html><head><meta http-equiv='Content-Type' content='text/html; charset=utf-8' /><meta name='ProgId' content='Word.Document'><meta name='Generator' content='Microsoft Word 15'><meta name='Originator' content='Microsoft Word 15'></head><body><div class='WordSection1'></div></body></html>";

export const useAction = () => {
  const [content, setContent] = useState<string>("");

  const [input, setInput] = useState<string>("");

  let timer;

  const [editor, setEditor] = useState<wangEditor.IDomEditor | null>(null);

  const [html, setHtml] = useState<string>("");

  const [showContent, setShowContent] = useState<string>("");

  const [num, setNum] = useState<number>(0);

  const [isOk, setIsOk] = useState<boolean>(false);

  const debouncedValue = useDebounce(html, { wait: 500 });

  const [language, setLanguage] = useState<string>("");

  const [isLoading, setIsLoading] = useState<boolean>(false);

  useEffect(() => {
    getBodyHtml();
  }, []);

  const getBodyHtml = () => {
    setIsLoading(true);
    Office.context.mailbox.item.saveAsync(function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const itemId = result.value;
        Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, async (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const token = result.value;

            // REST API
            const apiHtmlUrl = Office.context.mailbox.restUrl + "/v2.0/me/messages/" + itemId;
            await axios
              .get(apiHtmlUrl, {
                headers: {
                  Authorization: "Bearer " + token,
                },
              })
              .then(async (response) => {
                const bodyHtml = response.data.Body.Content;

                let isNew = false;

                if ("From" in response.data) {
                  isNew = false;
                } else {
                  isNew = true;
                }

                if (bodyHtml) {
                  const html = await getCleanHtml(ApiType.Rest, bodyHtml, isNew, itemId, token);
                  setShowContent(html);
                  setContent(html);
                  setIsLoading(false);
                } else {
                  // Office
                  Office.context.mailbox.item.body.getAsync("html", async (result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                      const html = await getCleanHtml(ApiType.Office, result.value, isNew);
                      setShowContent(html);
                      setContent(html);
                      setIsLoading(false);
                    }
                  });
                }
              });
          }
        });
      }
    });
  };

  // 清洗 html
  const getCleanHtml = async (type: ApiType, html: string, isNew: boolean, id?: string, token?: string) => {
    console.log(type, id, token);
    let parser = new DOMParser();
    const doc = parser.parseFromString(html, "text/html");
    const wordSection1Div = doc.querySelector(".WordSection1");

    const pElements = wordSection1Div.querySelectorAll("p");

    // 更新图片路径
    switch (type) {
      case ApiType.Office:
        doc.documentElement.innerHTML = doc.documentElement.innerHTML.replace(/>\s+</g, "><");
        break;
      case ApiType.Rest: {
        const images = wordSection1Div.querySelectorAll("img");

        if (images.length > 0) {
          const getAccachment = async (itemId: string) => {
            const attachments = [];
            const apiAttachmentsUrl = Office.context.mailbox.restUrl + "/v2.0/me/messages/" + itemId + "/attachments";
            await axios
              .get(apiAttachmentsUrl, {
                headers: {
                  Authorization: "Bearer " + token,
                },
              })
              .then((response) => {
                // 获取附件list
                response.data.value?.map((item) => {
                  item.IsInline &&
                    attachments.push({
                      id: item.Id,
                      contentId: item.ContentId,
                      base64: item.ContentBytes,
                    });
                });
                let index = 0;
                images.forEach((image) => {
                  if (image.src.includes("cid:")) index++;
                });
                if (index > 0 && attachments.length < index) {
                  timer = setInterval(async () => {
                    await getAccachment(itemId);
                  }, 2000);
                } else {
                  clearInterval(timer);

                  images.forEach((image) => {
                    if (image.src.includes("cid:")) {
                      const split = image.src.split("cid:")[1];
                      const index = attachments.findIndex((item) => item.contentId === split);
                      if (index > -1) {
                        image.src = "data:image/png;base64," + attachments[index].base64;
                      }
                    }
                  });
                }
              });
          };

          await getAccachment(id);
        }

        break;
      }

      default:
        break;
    }

    if (isNew) {
      pElements.forEach((item) => {
        const spanElements = item.querySelectorAll("span");

        for (let i = 0; i < spanElements.length; i++) {
          const spanElement = spanElements[i];
          const imgElements = spanElement.querySelectorAll("img");

          if (imgElements.length > 0) {
            // 将img标签移到span标签前面
            for (let j = 0; j < imgElements.length; j++) {
              const imgElement = imgElements[j];
              spanElement.parentNode.insertBefore(imgElement, spanElement);
            }

            // 移除span标签
            spanElement.parentNode.removeChild(spanElement);
          }
        }
      });
    } else {
      let hasMailOriginalLink = false;

      for (const item of Array.from(pElements)) {
        const aElement = item.querySelector("a[name='_MailOriginal']");
        if (aElement) {
          hasMailOriginalLink = true;
          break; // 终止循环
        }
      }

      if (hasMailOriginalLink) {
        // 清除需要放进富文本中p标签包含的span
        for (const item of Array.from(pElements)) {
          const aElement = item.querySelector("a[name='_MailOriginal']");
          if (aElement) {
            break; // 终止循环
          }
          const spanElements = item.querySelectorAll("span");

          for (let i = 0; i < spanElements.length; i++) {
            const spanElement = spanElements[i];
            const imgElements = spanElement.querySelectorAll("img");

            if (imgElements.length > 0) {
              // 将img标签移到span标签前面
              for (let j = 0; j < imgElements.length; j++) {
                const imgElement = imgElements[j];
                spanElement.parentNode.insertBefore(imgElement, spanElement);
              }

              // 移除span标签
              spanElement.parentNode.removeChild(spanElement);
            }
          }
        }
      } else {
        // 基础html标签
        const base = parser.parseFromString(baseHtml, "text/html");

        const headElement = base.getElementsByTagName("head")[0];

        const baseDivElement = base.querySelector(".WordSection1");

        const styleTags = doc.getElementsByTagName("style");
        for (let i = 0; i < styleTags.length; i++) {
          headElement.appendChild(styleTags[i]);
        }

        const parentElement = wordSection1Div.parentNode;

        let sameLevelDivs = Array.from(parentElement.children).filter((childElement) => childElement.tagName === "DIV");

        let originalDiv = "";

        let originalEmailInfo = "";

        for (let i = 0; i < sameLevelDivs.length; i++) {
          if (sameLevelDivs[i].classList.contains("WordSection1")) {
            originalDiv = sameLevelDivs[i].innerHTML;
          } else {
            let wrapperElement = base.createElement("p");
            wrapperElement.setAttribute("class", "MsoNormal");
            wrapperElement.innerHTML = `<a name='_MailOriginal'>${sameLevelDivs[i].innerHTML}</a>`;
            sameLevelDivs[i].innerHTML = wrapperElement.outerHTML;
            originalEmailInfo += sameLevelDivs[i].outerHTML;
          }
        }

        baseDivElement.innerHTML = originalEmailInfo + originalDiv;

        doc.documentElement.innerHTML = base.documentElement.innerHTML;
      }
    }

    parser = null;

    return doc.documentElement.outerHTML;
  };

  // 获取html到富文本框
  useEffect(() => {
    if (content) {
      let parser = new DOMParser();
      const doc = parser.parseFromString(content, "text/html");
      const wordSection1Div = doc.querySelector(".WordSection1");

      const aElement = wordSection1Div.querySelector("a[name='_MailOriginal']");

      const pElements = wordSection1Div.querySelectorAll("p");

      let pElementsOuterHTML = "";

      // 判断是否属于新建还是回复/转发
      if (aElement) {
        for (const item of Array.from(pElements)) {
          const aElement = item.querySelector("a[name='_MailOriginal']");
          if (aElement) {
            break; // 终止循环
          }
          pElementsOuterHTML += item.outerHTML;
        }
        setHtml(pElementsOuterHTML);
      } else {
        // 新建
        pElements.forEach((item) => {
          pElementsOuterHTML += item.outerHTML;
        });
        setHtml(pElementsOuterHTML);
      }

      parser = null;
    }
  }, [content]);

  useEffect(() => {
    if (debouncedValue) {
      setNum((prev) => prev + 1);
      if (isOk) {
        changeHtml(debouncedValue);
      }
    }
  }, [debouncedValue]);

  const changeHtml = async (debouncedValue: string) => {
    const segments = debouncedValue.match(/<p>(.*?)<\/p>/g) ?? [];

    let parser = new DOMParser();
    const doc = parser.parseFromString(content, "text/html");
    const wordSection1Div = doc.querySelector(".WordSection1");
    const paragraphs = wordSection1Div.querySelectorAll("p");
    const aElement = wordSection1Div.querySelector("a[name='_MailOriginal']");

    // 判断属于新建还是回复/转发
    if (aElement) {
      // 回复
      for (const paragraph of Array.from(paragraphs)) {
        const aElement = paragraph.querySelector("a[name='_MailOriginal']");
        if (aElement) {
          break; // 终止循环
        }
        paragraph.parentNode.removeChild(paragraph);
      }

      for (let i = segments.length - 1; i >= 0; i--) {
        const newParagraph = doc.createElement("p");
        newParagraph.setAttribute("class", "MsoNormal");
        newParagraph.innerHTML = await translate(
          segments[i].replace(/<p[^>]*>|<\/p>/g, "").replace(/"/g, "'"),
          isHTMLValid(segments[i].replace(/<p[^>]*>|<\/p>/g, "").replace(/"/g, "'"))
        );
        wordSection1Div.insertBefore(newParagraph, wordSection1Div.firstChild);
      }
    } else {
      paragraphs.forEach((paragraph) => {
        paragraph.parentNode.removeChild(paragraph);
      });

      for (let i = 0; i < segments.length; i++) {
        const newParagraph = doc.createElement("p");
        newParagraph.setAttribute("class", "MsoNormal");
        newParagraph.innerHTML = await translate(
          segments[i].replace(/<p[^>]*>|<\/p>/g, "").replace(/"/g, "'"),
          isHTMLValid(segments[i].replace(/<p[^>]*>|<\/p>/g, "").replace(/"/g, "'"))
        );
        wordSection1Div.appendChild(newParagraph);
      }
    }

    Office.context.mailbox.item.body.setAsync(doc.documentElement.outerHTML, { coercionType: "html" }, async () => {
      Office.context.mailbox.item.saveAsync(function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          setShowContent(doc.documentElement.outerHTML);
        } else {
          message.error("保存失败");
        }
      });
    });

    parser = null;
  };

  const translate = async (html: string, isHtml: boolean) => {
    const res = await PostTranslate(html, "zh-TW", isHtml);
    if (res.msg) {
      return html;
    }

    return res as string;
  };

  const isHTMLValid = (htmlText: string) => {
    var tagRegex = /<([a-z]+\d*)[\s\S]*>[\s\S]*?<\/\1>/gi;
    return tagRegex.test(htmlText);
  };

  // 跳过第一个赋值时重新调用setAsync方法
  useEffect(() => {
    if (num === 1) {
      setIsOk(true);
    }
  }, [num]);

  return { content, setContent, input, setInput, html, setHtml, showContent, editor, setEditor, isLoading };
};
