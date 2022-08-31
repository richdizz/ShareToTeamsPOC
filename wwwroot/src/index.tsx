import * as React from "react";
import * as ReactDOM from "react-dom";
import App from "./Components/App";
import { Provider, teamsTheme } from "@fluentui/react-northstar";

// Render the main component
ReactDOM.render(
    <Provider theme={teamsTheme}>
        <App />
    </Provider>,
    document.getElementById("page")
);