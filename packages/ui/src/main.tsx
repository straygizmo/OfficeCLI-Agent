import { createRoot } from "react-dom/client";
import { App } from "./App";
import { hostAppFromQuery } from "./host-client";

createRoot(document.getElementById("root")!).render(<App hostApp={hostAppFromQuery()} />);
