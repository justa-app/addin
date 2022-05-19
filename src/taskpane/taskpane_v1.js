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
    document.getElementById("search").onclick = run;
//    setInterval(function(){run()}, 1000)
  }
});

export async function run() {

        let profile = Office.context.mailbox.userProfile;

//        document.getElementById("item-subject").innerHTML = profile.accountType;
//        document.getElementById("item-subject").innerHTML = profile.displayName;
//        document.getElementById("item-subject").innerHTML = profile.emailAddress;
//        document.getElementById("item-subject").innerHTML = profile.timeZone;

        Office.context.mailbox.item.body.getAsync(
            "text",
            async function callback(result) {

                let text = result.value.toLowerCase();

                text = clean_punctuation(text);

                text = clean_line_breaks(text);

                text = clean_stop_worlds(text);

//                document.getElementById("item-subject").innerHTML = text;

                try {
                    let resource = `https://addin.dev.askjusta.com/search?q=${text}`
                    let init = {
                        'method': 'GET',
                        'headers': {
                                        'Content-Type': 'application/json',
                                   }
                    }

                    let response = await fetch(resource, init);
                    let items = await response.json();


                    let item_subject = document.getElementById("item-subject");

                    while (item_subject.hasChildNodes()) {
                      item_subject.removeChild(item_subject.firstChild);
                    }

                    let title = document.createElement('div');
                    title.innerHTML = 'We found somethings that might be useful'
                    title.style.cssText = `
                    color: #2955c1; font-size: 1.25rem; font-family: Helvetica,sans-serif; font-weight: 500;
                     line-height: 1.6; margin-bottom: 20px;
                    `
                    document.getElementById('item-subject').appendChild(title);

                    let ul = document.createElement('ul');
                    document.getElementById('item-subject').appendChild(ul);

                    items.forEach(item => {
                        let li = document.createElement('li');

                        let html_item = `
                        <a href=${item.url} style="margin-top: 0px; margin-bottom: 0px">${item.title}</a>
                        <h5 style="margin-top: 0px; margin-bottom: 0px">created_by: ${item.created_by}, type: ${item.type}</h5>
                        <h4 style="margin-top: 0px; margin-bottom: 20px">tags:  ${item.tags}</h4>
                        `
                        li.innerHTML = html_item;

                        ul.appendChild(li);
                    });


                    // summarizing the email

                    let email_analysing_paragraph = document.createElement('div');
                    email_analysing_paragraph.innerHTML = 'Email content analysed'
                    email_analysing_paragraph.style.cssText = `
                    color: #2955c1; font-size: 1.25rem; font-family: Helvetica,sans-serif; font-weight: 500;
                     line-height: 1.6; margin-bottom: 20px;
                    `
                    document.getElementById('item-subject').appendChild(email_analysing_paragraph)


                    // tags
                    resource = `https://addin.dev.askjusta.com/keywords`
                    init = {
                        'method': 'POST',
                        'headers': {
                                        'Content-Type': 'application/json',
                                   },
                        'body': result.value
                    }

                    response = await fetch(resource, init);
                    let tags = await response.json();

                    let tags_element = document.createElement('p');
                    tags_element.innerHTML = 'Tags: ' + JSON.stringify(tags)
                    document.getElementById('item-subject').appendChild(tags_element)

                    // summary
                    resource = `https://addin.dev.askjusta.com/summary`
                    init = {
                        'method': 'POST',
                        'headers': {
                                        'Content-Type': 'application/json',
                                   },
                        'body': result.value
                    }

                    response = await fetch(resource, init);
                    let summary = await response.json();

                    let summary_element = document.createElement('p');
                    summary_element.innerHTML = 'Summary: ' + summary
                    document.getElementById('item-subject').appendChild(summary_element)

                }
                catch (error) {
                    document.getElementById("item-subject").innerHTML = error;
                    return [["ERROR", error.message]];
                }
            });
}

export function clean_punctuation(rawString) {

    var rawLetters = rawString.split('');
    var punctuation = '!"#$%&\'()*+,-./:;<=>?@[\\]^_`{|}~';
    var rawLetters = rawString.split('');
    var cleanLetters = rawLetters.filter(function(letter) {
      return punctuation.indexOf(letter) === -1;
    });

    var cleanString = cleanLetters.join('');

    return cleanString
}

export function clean_line_breaks(rawString) {

   var cleanString = rawString.replace(/(\r\n|\n|\r)/gm," ");

   return cleanString
}

export function clean_stop_worlds(text) {

    let stop_worlds = ["send", "thanks", "please", "hi", "i", "me", "my", "myself", "we", "our", "ours", "ourselves",
     "you", "your", "yours", "yourself", "yourselves", "he", "him", "his", "himself", "she", "her", "hers", "herself",
     "it", "its", "itself", "they", "them", "their", "theirs", "themselves", "what", "which", "who", "whom", "this",
     "that", "these", "those", "am", "is", "are", "was", "were", "be", "been", "being", "have", "has", "had", "having",
     "do", "does", "did", "doing", "a", "an", "the", "and", "but", "if", "or", "because", "as", "until", "while", "of",
     "at", "by", "for", "with", "about", "against", "between", "into", "through", "during", "before", "after", "above",
     "below", "to", "from", "up", "down", "in", "out", "on", "off", "over", "under", "again", "further", "then", "once",
     "here", "there", "when", "where", "why", "how", "all", "any", "both", "each", "few", "more", "most", "other",
     "some", "such", "no", "nor", "not", "only", "own", "same", "so", "than", "too", "very", "s", "t", "can", "will",
      "just", "don", "should", "now"];

    text = text.split(' ').filter(item => !stop_worlds.includes(item));

    text = text.join(' ')

    return text
}

export async function send_to_suggestion() {
    Office.context.mailbox.item.to.setAsync( ['yohai@justa.app'] );
}

export function show_suggestion() {
    document.getElementById('to_suggestion').style.display = 'block';
}

export function hide_suggestion() {
    document.getElementById('to_suggestion').style.display = 'none';
}