import * as React from "react";
import { useAction } from "./hook";
import "../taskpane.css";
import { Select } from "antd";

const App = () => {
  const { content } = useAction();

  const handleChange = (value: string) => {
    console.log(`selected ${value}`);
  };

  return (
    <div className="w-full h-screen bg-red-300 flex flex-col overflow-auto">
      <div className="w-full bg-blue-300">
        <Select
          className="w-full"
          defaultValue="lucy"
          onChange={handleChange}
          options={[
            { value: "jack", label: "Jack" },
            { value: "lucy", label: "Lucy" },
            { value: "Yiminghe", label: "yiminghe" },
            { value: "disabled", label: "Disabled", disabled: true },
          ]}
        />
      </div>
      <div dangerouslySetInnerHTML={{ __html: content }} />
    </div>
  );
};

export default App;
