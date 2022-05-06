/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    setInterval(function(){run()}, 300)
  }
});

export async function run() {

    Office.context.mailbox.item.body.getAsync(
    "text",
    { asyncContext: "This is passed to the callback" },
    function callback(result) {
        if (result.value.toLowerCase().replace("  ", " ").includes('what does gdpr mean ?') == true) {
          document.getElementById("item-subject").innerHTML = gdpr_definition
        } else if (result.value.toLowerCase().replace("  ", " ").includes('what does soc2 mean ?') == true) {
          document.getElementById("item-subject").innerHTML = soc2_definition
        } else {
          document.getElementById("item-subject").innerHTML = 'analysing...'
        }
    });
}

const gdpr_definition = `
General Data Protection Regulation requirements prohibit companies from hiding behind illegible terms and
conditions that are difficult to understand. Instead, GDPR compliance requires companies to clearly define
their data privacy policies and make them easily accessible.
`
const soc2_definition = `
SOC 2 is a voluntary compliance standard for service organizations, developed by the American Institute of CPAs
 (AICPA), which specifies how organizations should manage customer data. The standard is based on the following
  Trust Services Criteria: security, availability, processing integrity, confidentiality, privacy.
`

