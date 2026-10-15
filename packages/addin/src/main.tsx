/* global Office */
import { createRoot } from "react-dom/client";
import { App } from "./App";
import { detectHostApp } from "./office-ops";

Office.onReady(() => {
  const root = createRoot(document.getElementById("root")!);
  try {
    root.render(<App hostApp={detectHostApp()} />);
  } catch (err) {
    root.render(<div className="msg system">{(err as Error).message}</div>);
  }
});
