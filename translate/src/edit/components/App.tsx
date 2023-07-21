import * as React from "react";
import { useAction } from "./hook";
import "../edit.css";
import { Editor, Toolbar } from "@wangeditor/editor-for-react";
import * as wangEditor from "@wangeditor/editor";
import "@wangeditor/editor/dist/css/style.css";
import { PostAttachmentUpload } from "../../api";

type InsertImageFnType = (url: string, alt: string, href: string) => void;

const App = () => {
  const { html, setHtml, showContent, editor, setEditor, isLoading } = useAction();

  const toolbarConfig: Partial<wangEditor.IToolbarConfig> = {
    toolbarKeys: [
      "fontFamily",
      "fontSize",
      "color",
      "|",
      "bold",
      "italic",
      "underline",
      "through",
      "bgColor",
      "sup",
      "sub",
      "|",
      "bulletedList",
      "numberedList",
      "justifyJustify",
      "delIndent",
      "indent",
      "|",
      "insertLink",
      "redo",
      "undo",
      "uploadImage",
    ],
  };

  const editorConfig = {
    autoFocus: false,
    hoverbarKeys: {
      attachment: {
        menuKeys: ["downloadAttachment"], // “下载附件”菜单
      },
    },
    MENU_CONF: {
      uploadImage: {
        async customUpload(file: File, insertFn: InsertImageFnType) {
          const formData = new FormData();
          formData.append("file", file);
          PostAttachmentUpload(formData).then((res) => {
            if (res && res.data) {
              insertFn(res.data.fileUrl, res.data.fileName, res.data.filePath);
            }
          });
        },
      },
    },
  };

  return (
    <div className="w-full h-screen flex flex-col relative">
      {isLoading && <div className="absolute w-screen h-screen z-10 bg-black opacity-10" />}
      <div dangerouslySetInnerHTML={{ __html: showContent }} className="flex-1 no-scrollbar overflow-y-auto" />
      <div className="h-[28rem] border border-solid border-gray-300">
        <Toolbar
          editor={editor}
          defaultConfig={toolbarConfig}
          mode="default"
          style={{ borderBottom: "1px solid #ccc" }}
        />
        <Editor
          defaultConfig={editorConfig}
          value={html}
          onCreated={setEditor}
          onChange={(editor) => {
            setHtml(editor.getHtml());
          }}
          mode="default"
          className="flex-1 h-[20rem]"
        />
      </div>
    </div>
  );
};

export default App;
