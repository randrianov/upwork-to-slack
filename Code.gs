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

  let label = GmailApp.getUserLabelByName("upwork");
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
      formattedBody = convertLinksToSlackFormat(cutEmailBody(emailBody));

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
    });

    //thread.removeLabel(label);
  });
}

function cutEmailBody(emailBody) {
  let topMessage = extractTopMessage(emailBody);
  let cutoffText = "Reply: https://www.upwork.com";
  let index = topMessage.indexOf(cutoffText);

  return index !== -1 ? topMessage.substring(0, index).trim() : topMessage;
}

function extractTopMessage(emailBody) {
  let normalizedBody = emailBody.replace(/\n+/g, " ").replace(/\s+/g, " ");
  
  let separatorPattern = /On \w{3}, \d{1,2} \w{3} \d{4} at /i; // reply chain starts with "On Fri, 7 Feb 2025 at 20:15"
  let match = normalizedBody.match(separatorPattern);

  return match ? emailBody.substring(0, match.index).trim() : emailBody.trim();
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
  Logger.log("Response: " + json); // Log API response
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
