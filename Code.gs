const SLACK_TOKEN = 'xoxb-xxxxx'; // Slack Bot Token
const USER_TOKEN = "xoxp-xxxxx"; // For searching Slack messages
const FWR_PREFIX = "Fwd: ";
const CHANNEL = 'CXXXXXXXXXX'; // Slack channel ID
const CHANNEL_NAME = "XXX";
const RECIPIENT = "upwork-notifications-xxxxx@xxxxx-team.slack.com"; // Slack channel email

function sendEmailToSlack() {
  let threads = GmailApp.search("label:upwork is:unread");
  if (threads.length === 0) {
    Logger.log("No unread emails found.");
    return;
  }

  let subject = "";
  let threadTs = null;
  let isMessageForwarded = false;
  let emailBody = "";
  let formattedBody = "";

  threads.forEach(thread => {
    let unreadMessages = thread.getMessages().filter(msg => msg.isUnread() && !msg.isDraft() && !msg.isInTrash());

    Logger.log(`Found ${unreadMessages.length} unread messages in thread: ${thread.getFirstMessageSubject()}`);
    
    subject = thread.getFirstMessageSubject();

    unreadMessages.forEach((message, index) => {
      emailBody = message.getPlainBody();

      formattedBody = cutEmailBody(emailBody);
      formattedBody = removeLinks(formattedBody);
      formattedBody = normalizeNewlines(formattedBody);
      //formattedBody = convertLinksToSlackFormat(cutEmailBody(emailBody));
      Logger.log(`formattedBody ${formattedBody}`);
      //Logger.log(`getCharCodes ${getCharCodes(formattedBody)}`);

      if (isDuplicateMessage(formattedBody)) {
        Logger.log(`Duplicate message detected. Skipping... ${formattedBody}`);
        message.markRead();
        return; // Skips to the next iteration
      }

      if (index === 0) {
        threadTs = findSlackMessage(FWR_PREFIX + subject);
      }

      if (isMessageForwarded && threadTs === null) {
        threadTs = findSlackMessage(FWR_PREFIX + subject);
        // try for up to 180 sec to get threadTs
        for (let i = 0; i < 36 && !threadTs; i++) {
          Utilities.sleep(5000); 
          threadTs = findSlackMessage(FWR_PREFIX + subject);
        }
      }

      if (threadTs) {
        postToSlack(formattedBody, threadTs);
      } else {
        autoForwardMessage(message.getId());
        Logger.log("Email sent to: " + RECIPIENT);
        isMessageForwarded = true;

        Utilities.sleep(5000); // Wait 5 seconds for Gmail to register the sent email
        let sentThreads = GmailApp.search(`from:me to:${RECIPIENT} subject:${FWR_PREFIX}${subject}`);
        
        if (sentThreads.length > 0) {
          let lastSentThread = sentThreads[sentThreads.length - 1];
          Logger.log("Forwarded email removed from Sent folder: " + lastSentThread.getFirstMessageSubject());
          deleteThreadViaAPI(lastSentThread.getId());
        }
      }

      message.markRead();
      storeMessageHash(formattedBody);
    });
  });
}

function cutEmailBody(emailBody) {
  let topMessage = extractTopMessage(emailBody);
  let cutoffTextHeader1 = "sent a message";
  let cutoffTextHeader2 = "Unread message from"
  let cutoffTextFooter1 = "View on Upwork";
  let cutoffTextFooter2 = "Reply: https://www.upwork.com";
  let index;

  index = topMessage.indexOf(cutoffTextHeader1);
  if (index == -1) {
    index = topMessage.indexOf(cutoffTextHeader2);
  }

  if (index !== -1) {
    // cut off header
    topMessage = topMessage.substring(index).trim();
  }
  
  index = topMessage.indexOf(cutoffTextFooter1);
  if (index == -1) {
    index = topMessage.indexOf(cutoffTextFooter2);
  }
  // cut off footer
  return index !== -1 ? topMessage.substring(0, index).trim() : topMessage;
}

function extractTopMessage(emailBody) {
  let separatorPattern = /On \w{3}, \d{1,2} \w{3,4} \d{4} at /i; // reply chain starts with "On Fri, 7 Feb 2025 at 20:15"
  let match = emailBody.match(separatorPattern);

  return match ? emailBody.substring(0, match.index).trim() : emailBody.trim();
}

function removeLinks(text) {
  return text.replace(/(https?:\/\/[^\s]+)/g, "");
}

function convertLinksToSlackFormat(text) {
  return text.replace(/(https?:\/\/[^\s]+)/g, "<$1|🔗 Link>");
}

function findSlackMessage(query) {
  let searchQuery = encodeURIComponent(`in:#${CHANNEL_NAME} ${query}`);
  let url = `https://slack.com/api/search.messages?query=${searchQuery}&count=1`;

  let response = UrlFetchApp.fetch(url, {
    method: "get",
    headers: { "Authorization": `Bearer ${USER_TOKEN}` },
    muteHttpExceptions: true
  });

  let json = JSON.parse(response.getContentText());
  return json.ok && json.messages.total > 0 ? json.messages.matches[0].ts : null;
}

function postToSlack(text, threadTs) {
  let payload = {
    channel: CHANNEL,
    text: text,
    thread_ts: threadTs
  };

  let options = {
    method: "post",
    contentType: "application/json",
    headers: { "Authorization": `Bearer ${SLACK_TOKEN}` },
    payload: JSON.stringify(payload)
  };

  UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', options);
  Logger.log("Payload sent: " + text);
}

function autoForwardMessage(messageId) {
  let message = GmailApp.getMessageById(messageId);
  let subject = message.getSubject();
  let body = message.getBody();

  let email = `To: ${RECIPIENT}\r\n` +
              `Subject: ${FWR_PREFIX}${subject}\r\n` +
              `MIME-Version: 1.0\r\n` +
              `Content-Type: text/html; charset=UTF-8\r\n` +
              `X-Google-Original-From: ${Session.getActiveUser().getEmail()}\r\n\r\n` + 
              body;

  let encodedEmail = Utilities.base64EncodeWebSafe(email);

  let options = {
    method: "post",
    contentType: "application/json",
    headers: { "Authorization": `Bearer ${ScriptApp.getOAuthToken()}` },
    payload: JSON.stringify({ raw: encodedEmail }),
    muteHttpExceptions: true
  };

  let response = UrlFetchApp.fetch("https://www.googleapis.com/gmail/v1/users/me/messages/send", options);
  Logger.log(response.getContentText());
}

function deleteThreadViaAPI(threadId) {
  let url = `https://www.googleapis.com/gmail/v1/users/me/threads/${threadId}`;

  let options = {
    method: "delete",
    headers: { "Authorization": `Bearer ${ScriptApp.getOAuthToken()}` },
    muteHttpExceptions: true
  };

  let response = UrlFetchApp.fetch(url, options);
  Logger.log("Deleted Thread ID: " + threadId + " Response: " + response.getResponseCode());
}

function generateMessageHash(formattedBody) {
    let withoutTimeFormattedBody = cutTimeFromBodyForHashing(formattedBody);
    let noRoomLinkWithoutTimeFormattedBody = cutRoomLink(withoutTimeFormattedBody);
    Logger.log(`🕒 Removed timestamp: ${withoutTimeFormattedBody}`);
    return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, noRoomLinkWithoutTimeFormattedBody));
}

function isDuplicateMessage(body) {
    Logger.log(`\n🔍 isDuplicateMessage message`);
    Logger.log(`   Body (clipped): ${body}`);

    var cache = CacheService.getScriptCache();
    var messageHash = generateMessageHash(body);
    Logger.log(`💾 Storing message hash: ${messageHash}`);
    return cache.get(messageHash) !== null; // If hash exists, it's a duplicate
}

function storeMessageHash(body) {
    var cache = CacheService.getScriptCache();
    var messageHash = generateMessageHash(body);
    cache.put(messageHash, "1", 3600); // Store hash for 1 hour (3600 seconds)
}


function cutTimeFromBodyForHashing(body) {
    var dateTimePattern = /\b\d{1,2}:\d{2}\s*(AM|PM)?\s*[A-Z]*,\s*\d{1,2}\s*[A-Za-z]+\s*\d{4}\b/;
    // convert 8:09 AM EET, 28 Mar 2025 → 28 Mar 2025, cause same message could have different time (gmail local settings)
    // Remove only the first match
    var normalizedBody = body.replace(dateTimePattern, (match) => {
        Logger.log(`🕒 Removing timestamp: ${match}`);
        return match.replace(/\b\d{1,2}:\d{2}\s*(AM|PM)?\s*[A-Z]*,\s*/, ""); // Keep only the date
    });

    return normalizedBody;
}

// remove room url params after (?) cause they are different depends on user <https://www.upwork.com/ab/messages/rooms/room_ce05a0abc18250a1589ae2fe32b7eb99?companyReference=424277385192652801&app_type=fl&frkscc=qnbKjchprKoz|🔗 Link>
function cutRoomLink(text) {
    return text.replace(/<https:\/\/www\.upwork\.com\/ab\/messages\/rooms\/([^?|>]+)[^>]*\|🔗 Link>/g, (match, roomId) => {
        const cleanedUrl = `https://www.upwork.com/ab/messages/rooms/${roomId}`;
        Logger.log(`🔗 Converting Upwork link: ${match} → <${cleanedUrl}>`);
        return `<${cleanedUrl}>`;
    });
}

function normalizeNewlines(text) {
  // Replace all variations of multiple CR/LF into a single \n
  return text.replace(/(\r\n|\r|\n){2,}/g, '\n');
}

function getCharCodes(text) {
  let charCodeString = '';
  for(let i = 0; i < text.length; i++){
    let code = text.charCodeAt(i);
    charCodeString += code + ' ';
  }
  return charCodeString;
}
