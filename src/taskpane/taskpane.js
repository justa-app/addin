/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("to_suggestion").onclick = send_to_suggestion;
    setInterval(function(){run()}, 1000)
  }
});

export async function run() {

    Office.context.mailbox.item.body.getAsync(
    "text",
    { asyncContext: "This is passed to the callback" },
    function callback(result) {
        if (result.value.toLowerCase().replace("  ", " ").includes('gdpr') == true) {
          document.getElementById("item-subject").innerHTML = gdpr_definition
          show_suggestion()
          let html_body = result.value.replace('gdpr', '<span style="background-color:#00FEFE">gdpr</span>')
          let body = '<div>' + html_body + '</div>'
          body = body.replace(/\s+/g, ' ').trim();
          Office.context.mailbox.item.body.setAsync(
                                                            body,
                                                            {coercionType: Office.CoercionType.Html});


        } else if (result.value.toLowerCase().replace("  ", " ").includes('soc2') == true) {
          document.getElementById("item-subject").innerHTML = soc2_definition
          show_suggestion()
          let html_body = result.value.replace('soc2', '<span style="background-color:#00FEFE">soc2</span>')
          let body = '<div>' + html_body + '</div>'
          body = body.replace(/\s+/g, ' ').trim();
          Office.context.mailbox.item.body.setAsync(
                                                  body,
                                                  {coercionType: Office.CoercionType.Html});


        } else if (result.value.toLowerCase().replace("  ", " ").includes('hipaa') == true) {
          document.getElementById("item-subject").innerHTML = hipaa_definition
          show_suggestion()

          let html_body = result.value.replace('hipaa', '<span style="background-color:#00FEFE">hipaa</span>')
          let body = '<div>' + html_body + '</div>'

          body = body.replace(/\s+/g, ' ').trim();

          Office.context.mailbox.item.body.setAsync(
                                        body,
                                        {coercionType: Office.CoercionType.Html});
        }
        else {
          if (result.value.length == 1){
                document.getElementById("item-subject").innerHTML = ''
          } else {
                document.getElementById("item-subject").innerHTML = 'analysing...'
          }
          hide_suggestion()
        }
    });
}

export async function send_to_suggestion() {
    Office.context.mailbox.item.to.setAsync( ['oren@justa.app'] );
}

export function show_suggestion() {
    document.getElementById('to_suggestion').style.display = 'block';
}

export function hide_suggestion() {
    document.getElementById('to_suggestion').style.display = 'none';
}

const gdpr_definition = `
<h1>Justa.app analysis</h1>
<h2>General Information</h2>
General Data Protection Regulation requirements prohibit companies from hiding behind illegible terms and
conditions that are difficult to understand. Instead, GDPR compliance requires companies to clearly define
their data privacy policies and make them easily accessible.

<h2>Usable resources</h2>
We also found for you some relevant and usable information:
   <ul>
    <li><a href="https://gdpr.eu/">Guidance</a></li>
    <li><a href="https://gdpr-info.eu/">Insight</a></li>
  </ul>
`

const soc2_definition = `
<h1>Justa.app analysis</h1>
<h2>General Information</h2>
SOC 2 is a voluntary compliance standard for service organizations, developed by the American Institute of CPAs
 (AICPA), which specifies how organizations should manage customer data. The standard is based on the following
  Trust Services Criteria: security, availability, processing integrity, confidentiality, privacy.

<h2>Usable resources</h2>
  We also found for you some relevant and usable information:
   <ul>
    <li><a href="https://www.imperva.com/learn/data-security/soc-2-compliance/">SOC 2 Compliance</a></li>
    <li><a href="https://www.checkpoint.com/cyber-hub/cyber-security/what-is-soc-2-compliance/">What is SOC 2 Compliance?</a></li>
    <li><a href="https://www.nextep.co.il/soc2/">Who needs?</a></li>
  </ul>
`

const hipaa_definition = `
<h1>Justa.app analysis</h1>
<h2>General Information</h2>
HIPAA stands for Health Insurance Portability and Accountability Act.
HIPAA Compliance is the process by which covered entities need to protect and secure
a patient's healthcare data or Protected Health Information.

<h2>Usable resources</h2>
We also found for you some relevant and usable information:
   <ul>
    <li><a href="https://www.hhs.gov/hipaa/index.html">Health information privacy</a></li>
  </ul>
`
