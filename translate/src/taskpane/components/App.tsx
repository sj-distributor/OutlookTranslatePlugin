import * as React from "react";
import { useAction } from "./hook";
import "../taskpane.css";
import { Button, Select } from "antd";

const App = () => {
  const { content, translateContent, handleChange, language } = useAction();

  return (
    <div className="w-full h-screen flex flex-col overflow-auto">
      <div className="w-full flex flex-row">
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
        <Button onClick={translateContent}>translate</Button>
      </div>

      <div dangerouslySetInnerHTML={{ __html: content }} />
    </div>
  );
};

export default App;
