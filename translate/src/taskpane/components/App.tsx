import * as React from "react";
import { useAction } from "./hook";
import "../taskpane.css";

const App = () => {
  const { content } = useAction();

  return (
    <div className="w-full h-screen flex flex-col overflow-auto p-1">
      <div dangerouslySetInnerHTML={{ __html: content }} />
    </div>
  );
};

export default App;
