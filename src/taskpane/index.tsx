import App from "./components/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import { ThemeProvider } from "@fluentui/react";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { AuthProvider } from "./components/AuthContext";

/* global document, Office, module, require */

initializeIcons();

let isOfficeInitialized = false;

const title = "Contoso Task Pane Add-in";

const render = (Component) => {
  ReactDOM.render(
    <AuthProvider>
      <AppContainer>
        <ThemeProvider>
          <Component title={title} isOfficeInitialized={isOfficeInitialized} />
        </ThemeProvider>
      </AppContainer>
    </AuthProvider>,
    document.getElementById("container")
  );
};

/* Render application after Office initializes */
Office.onReady(() => {
  isOfficeInitialized = true;
  render(App);
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}