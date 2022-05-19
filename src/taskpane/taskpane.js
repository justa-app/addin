Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("username").innerHTML = `Welcome ${Office.context.mailbox.userProfile.displayName}`
//    document.getElementById("search_full").onclick = full;
//    document.getElementById("search_tags").onclick = tags;
//    document.getElementById("search_summary").onclick = summary;
    setInterval(function(){summary()}, 5000)
  }
});

export async function full() {
    Office.context.mailbox.item.body.getAsync(
                "text",
                async function callback(result) {
                    let text = result.value.toLowerCase();
                    await knowledge_pieces_search(text);
                    document.getElementById("debug_message").innerHTML = `querying engine ... $(text)`;
                });
}

export async function tags() {
    Office.context.mailbox.item.body.getAsync(
                    "text",
                    async function callback(result) {
                        let text = result.value.toLowerCase();
                          let tags_res = await generate_tags(text);
                          await knowledge_pieces_search(tags_res.join(' '))
                          document.getElementById("debug_message").innerHTML = tags_res.join(' ');
                    });
}

export async function summary() {
    Office.context.mailbox.item.body.getAsync(
                        "text",
                        async function callback(result) {
                            let text = result.value.toLowerCase();
                            let summary_res = await generate_summary(text)
                            let tags_res = await generate_tags(summary_res);
                            await knowledge_pieces_search(tags_res.join(' '))
                            let yy = await store_tags(tags_res);
                            document.getElementById("debug_message").innerHTML = `querying engine ... (${tags_res.join(' ')})`;
                        });
}

async function knowledge_pieces_search(text) {
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

        let knowledge_pieces_div = document.getElementById("knowledge_pieces");

        while (knowledge_pieces_div.hasChildNodes()) {
          knowledge_pieces_div.removeChild(knowledge_pieces_div.firstChild);
        }

        let title = document.createElement('div');
        title.innerHTML = 'We found somethings that might be useful'
        title.style.cssText = `
        color: #2955c1; font-size: 1.25rem; font-family: Helvetica,sans-serif; font-weight: 500;
         line-height: 1.6; margin-bottom: 20px;
        `
        knowledge_pieces_div.appendChild(title);

        let ul = document.createElement('ul');
        knowledge_pieces_div.appendChild(ul);

        items.forEach(item => {
            let li = document.createElement('li');

            let html_item = `
            <a href=${item.url} style="margin-top: 0px; margin-bottom: 0px">${item.title}</a>
            <h5 style="margin-top: 0px; margin-bottom: 20px">created_by: ${item.created_by}, type: ${item.type}</h5>
            `
            li.innerHTML = html_item;

            ul.appendChild(li);
        });
    }
    catch (error) {
        knowledge_pieces_div.innerHTML = error;
//        return [["ERROR", error.message]];
    }
}

async function generate_summary(text) {
    let resource = `https://addin.dev.askjusta.com/summary`
    let init = {
        'method': 'POST',
        'headers': {
                        'Content-Type': 'application/json',
                   },
        'body': text
    }

    let response = await fetch(resource, init);
    let summary = await response.json();

    return summary
}

async function generate_tags(text) {
    let resource = `https://addin.dev.askjusta.com/keywords`
    let init = {
        'method': 'POST',
        'headers': {
                        'Content-Type': 'application/json',
                   },
        'body': text
    }

    let response = await fetch(resource, init);
    let tags = await response.json();

    return tags
}

async function store_tags(tags) {
    let resource = `https://addin.dev.askjusta.com/save`
    let init = {
        'method': 'POST',
        'headers': {
                        'Content-Type': 'application/json',
                   },
        'body': JSON.stringify({
            'user_id': Office.context.mailbox.userProfile.emailAddress,
            'email_id': Office.context.mailbox.item.conversationId,
            'tags': tags
        })
    }

    await fetch(resource, init);
}