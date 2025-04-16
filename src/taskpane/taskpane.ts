/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { App } from './components/App';
import { initializeIcons } from '@fluentui/react';

Office.onReady((info) => {
  console.log("Office.onReady triggered", info);
  if (info.host === Office.HostType.Excel) {
    try {
      // Initialize FluentUI icons
      initializeIcons();
      console.log("FluentUI icons initialized");
      
      // Hide loading information
      document.getElementById("sideload-msg").style.display = "none";
      document.getElementById("app-body").style.display = "flex";
      console.log("Display elements updated");
      
      // Render the React app
      ReactDOM.render(
        React.createElement(App),
        document.getElementById("app-root")
      );
      console.log("React app rendering completed");
    } catch (error) {
      console.error("Error in Office.onReady:", error);
    }
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}
