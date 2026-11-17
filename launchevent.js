/*
 * QTP Email Signature - Outlook event-based add-in
 * Inserts the standard QTP signature on new emails, replies and forwards,
 * turns off the user's own Outlook signature, and stamps a header so the
 * Exchange transport rule can skip messages that already have a signature.
 */

console.log("QTP Signature: script started");

// ---------------------------------------------------------------------------
// 1. STAFF DETAILS
//    Email address must be lower case. Leave a field as "" to hide that line.
// ---------------------------------------------------------------------------
var STAFF = {
  "eric.leonard@qtponline.com": {
    firstName: "Eric",
    lastName: "Leonard",
    title: "I.T Admin",
    ext: "1717",
    phone: "+353 44 000 0000",
    mobile: "+353 87 000 0000"
  }
};

// ---------------------------------------------------------------------------
// 2. SIGNATURE TEMPLATE (matches the current transport rule design)
// ---------------------------------------------------------------------------
var LOGO_URL = "https://qtp.aws.aphix.software/sites/admin/plugins/elfinder/files/qtp/logos/IT/QTP%20Logo%20Circle.png";
var SIGNATURE_HEADER = "X-QTP-Signature";

function esc(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function row(content, style) {
  return '<tr><td style="font-size:12pt;' + (style || "") + '">' + content + "</td></tr>";
}

function buildSignature(p, email) {
  var spacer = '<table cellpadding="0" cellspacing="0" border="0"><tr><td style="height:20px; line-height:20px; font-size:1px;">&nbsp;</td></tr></table>';
  var rows = "";

  rows += row(esc(p.firstName) + " " + esc(p.lastName), " font-weight:bold; color:#1f3f95;");
  if (p.title) rows += row(esc(p.title), " font-weight:bold; color:#f4a11a;");
  rows += row("Quality Tractor Parts", " color:#1f3f95; font-weight:bold;");
  rows += row("Units 5-7, Zone C, Mullingar Business Park, Mullingar, Co. Westmeath, Ireland. N91 ND3T", " color:#555555;");
  rows += row('<strong>Email:</strong> <a href="mailto:' + esc(email) + '" style="color:#333333; text-decoration:none;">' + esc(email) + "</a>");

  var phoneParts = [];
  if (p.ext) phoneParts.push("<strong>Ext.:</strong> " + esc(p.ext));
  if (p.phone) phoneParts.push("<strong>Tel:</strong> " + esc(p.phone));
  if (phoneParts.length) rows += row(phoneParts.join(" | "));
  if (p.mobile) rows += row("<strong>Mob:</strong> " + esc(p.mobile));

  rows += row('<a href="https://www.qtponline.com/" style="color:#1f3f95; text-decoration:none; font-weight:bold;">qtponline.com</a>');

  return (
    spacer +
    '<div style="font-family:Segoe UI, Arial, sans-serif; font-size:12pt; color:#333333;">Best Regards,</div>' +
    spacer +
    '<table cellpadding="0" cellspacing="0" border="0" style="font-family:Segoe UI, Arial, sans-serif; font-size:13px; color:#333333;"><tr>' +
      '<td style="padding-right:18px; vertical-align:middle;">' +
        '<a href="https://www.qtponline.com/"><img src="' + LOGO_URL + '" alt="Quality Tractor Parts" width="110" style="display:block; border:0;"></a>' +
      "</td>" +
      '<td style="border-left:3px solid #1f3f95; padding-left:18px; vertical-align:middle;">' +
        '<table cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">' + rows + "</table>" +
      "</td>" +
    "</tr></table>" +
    '<table cellpadding="0" cellspacing="0" border="0"><tr><td style="height:16px; line-height:20px; font-size:1px;">&nbsp;</td></tr></table>' +
    '<table cellpadding="0" cellspacing="0" border="0"><tr>' +
      '<td style="font-size:9pt; color:#888888; font-family:Aptos, Arial, sans-serif; border-top:1px solid #e0e0e0; padding-top:8px; line-height:1.15;">' +
        "This email and any attachments are for authorised use by the intended recipient(s) only. It may contain proprietary material, confidential information and/or be subject to legal privilege. It should not be copied, disclosed to, retained or used by any other party. If you are not an intended recipient then please promptly delete this email and any attachments and all copies, and inform the sender. Thank you.<br>" +
        "Original equipment manufacturer names and part numbers are quoted for reference purposes only and are not intended to infer that our replacement parts are used as original equipment. Please consider the environment before printing this email." +
      "</td>" +
    "</tr></table>"
  );
}

// ---------------------------------------------------------------------------
// 3. EVENT HANDLER - runs on new message, reply, reply all and forward
//    DEBUG VERSION: logs every step to the browser console and shows a
//    message bar in the compose window with the result.
// ---------------------------------------------------------------------------
var DEBUG = true;

function log(text, extra) {
  if (!DEBUG) return;
  if (extra !== undefined) {
    console.log("QTP Signature: " + text, extra);
  } else {
    console.log("QTP Signature: " + text);
  }
}

function note(item, text) {
  if (!DEBUG) return;
  try {
    item.notificationMessages.replaceAsync("qtpDebug", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: ("QTP Signature: " + text).substring(0, 150),
      icon: "Icon.16x16",
      persistent: false
    }, function (r) {
      if (r.status === Office.AsyncResultStatus.Failed) {
        log("message bar failed - " + r.error.message);
      }
    });
  } catch (e) {
    console.error("QTP Signature: message bar error", e);
  }
}

function onNewMessageComposeHandler(event) {
  log("handler called");
  var item;
  try {
    var mailbox = Office.context.mailbox;
    log("mailbox found", !!mailbox);

    item = mailbox.item;
    log("item found", !!item);

    var email = (mailbox.userProfile.emailAddress || "").toLowerCase();
    log("email address is " + email);

    var person = STAFF[email];
    if (!person) {
      log("no staff entry for " + email);
      note(item, "ran, but no staff entry for " + email);
      event.completed();
      return;
    }
    log("staff entry found");

    var html = buildSignature(person, email);
    log("signature built, length " + html.length);

    var insertSignature = function () {
      log("inserting signature");
      item.body.setSignatureAsync(html, { coercionType: Office.CoercionType.Html }, function (r) {
        if (r.status === Office.AsyncResultStatus.Failed) {
          console.error("QTP Signature: setSignature failed", r.error);
          note(item, "signature failed - " + r.error.message);
        } else {
          log("signature added");
          note(item, "signature added for " + email);
        }
        event.completed();
        log("event completed");
      });
    };

    var setHeader = function (r) {
      if (r && r.status === Office.AsyncResultStatus.Failed) {
        log("disable client signature failed - " + r.error.message);
      } else {
        log("client signature disabled (or skipped)");
      }

      if (item.internetHeaders && item.internetHeaders.setAsync) {
        var headers = {};
        headers[SIGNATURE_HEADER] = "Applied";
        log("setting header");
        item.internetHeaders.setAsync(headers, function (hr) {
          if (hr.status === Office.AsyncResultStatus.Failed) {
            log("header failed - " + hr.error.message);
          } else {
            log("header set");
          }
          insertSignature();
        });
      } else {
        log("internetHeaders not supported, skipping header");
        insertSignature();
      }
    };

    if (item.disableClientSignatureAsync) {
      log("disabling client signature");
      item.disableClientSignatureAsync(setHeader);
    } else {
      log("disableClientSignatureAsync not supported, skipping");
      setHeader();
    }
  } catch (e) {
    console.error("QTP Signature: error", e);
    if (item) note(item, "error - " + e.message);
    event.completed();
  }
}

// ---------------------------------------------------------------------------
// 4. REGISTER THE HANDLER WITH OUTLOOK
// ---------------------------------------------------------------------------
function registerHandler(source) {
  try {
    Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
    log("handler registered (" + source + ")");
  } catch (e) {
    console.error("QTP Signature: could not register handler", e);
  }
}

// Register straight away (needed for classic Outlook)
registerHandler("immediate");

// Register again once Office has finished starting up (helps Outlook on the web)
if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(function (info) {
    log("Office ready", info);
    registerHandler("onReady");
  });
}

log("script loaded");
