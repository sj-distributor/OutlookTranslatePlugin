import { Button } from "antd";
import * as React from "react";
import "../edit.css";
import { useAction } from "./hook";

const App = () => {
  const { content, type, translate, src, list, s, d, z } = useAction();

  return (
    <div className="w-full">
      <Button
        onClick={() => {
          Office.context.mailbox.item.body.setAsync(content, { coercionType: type }, function callback() {});
        }}
      >
        1
      </Button>

      <Button onClick={translate}>2</Button>

      {/* <div dangerouslySetInnerHTML={{ __html: content }} /> */}
      <div>{content}</div>

      {/* {list.map((item, index) => (
        <img src={item} alt={`${index}`} key={index} />
      ))} */}

      {/* {s.map((item, index) => (
        <div key={index}>{item}</div>
      ))} */}
      {d.map((item, index) => (
        <div key={index}>{item.name}</div>
      ))}

      {z.map((item, index) => (
        <div key={index}>{item}</div>
      ))}
    </div>
  );
};

export default App;
