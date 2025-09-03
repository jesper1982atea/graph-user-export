import React from "react";
import ReactDOM from "react-dom/client";
import MainGraphMeetingsAndChats from "./components/msgraph/GraphMeetingsAndChats";
import { HashRouter } from "react-router-dom";
import "./styles.css";

const root = ReactDOM.createRoot(document.getElementById("root"));
root.render(
  <React.StrictMode>
  <HashRouter>
      <MainGraphMeetingsAndChats />
  </HashRouter>
  </React.StrictMode>
);
