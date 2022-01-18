import React from "react";
import GraphClient from "./GraphClient";
// import { Welcome } from "./sample/Welcome";

// var showFunction = Boolean(process.env.REACT_APP_FUNC_NAME);

export default function Tab() {
  return (
    <div>
      <GraphClient />
      {/* <Welcome showFunction={showFunction} /> */}
    </div>
  );
}