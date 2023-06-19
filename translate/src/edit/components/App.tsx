import { Button, Select } from "antd";
import * as React from "react";
import "../edit.css";
import { useAction } from "./hook";

const App = () => {
  const { content, language, handleChange, translate } = useAction();

  return (
    <div className="w-full">
      <div className="flex flex-row">
        <Select
          className="w-full"
          value={language}
          onChange={handleChange}
          options={[
            { value: "zh-TW", label: "zh-TW" },
            { value: "en", label: "en" },
            { value: "es", label: "es" },
          ]}
        />
        <Button
          onClick={() => {
            translate();
          }}
        >
          change
        </Button>
      </div>

      <div dangerouslySetInnerHTML={{ __html: content }} />
    </div>
  );
};

export default App;
