/*
 * QTP Email Signature - Outlook event-based add-in (Microsoft Graph version)
 *
 * On new emails, replies and forwards:
 *   1. Signs in silently as the current Outlook user (nested app authentication).
 *   2. Reads ONLY that user's own profile from Entra ID via Microsoft Graph (User.Read).
 *   3. Builds the QTP signature from those fields and inserts it under the cursor.
 *   4. Turns off the user's own Outlook signature and stamps X-QTP-Signature.
 *
 * Nothing about staff is stored in this file. If anything fails, the add-in does
 * nothing and the Exchange transport rule adds the signature as before.
 *
 * Settings (Client ID, Tenant ID, logo, debug) are in build/config.js.
 * GitHub Actions builds this file plus Microsoft MSAL into launchevent.js.
 */

import { createNestablePublicClientApplication } from "@azure/msal-browser";

/* global Office, QTP_CONFIG */

var CONFIG = (typeof QTP_CONFIG !== "undefined" && QTP_CONFIG) || {};
var DEBUG = !!CONFIG.debug;
var ALLOWED_DOMAIN = "@qtponline.com";
var SIGNATURE_HEADER = "X-QTP-Signature";
var GRAPH_URL =
  "https://graph.microsoft.com/v1.0/me?$select=givenName,surname,displayName,jobTitle,businessPhones,mobilePhone,faxNumber,mail";

// ---------------------------------------------------------------------------
// Logging (only when debug is on)
// ---------------------------------------------------------------------------
function log(text, extra) {
  if (!DEBUG) return;
  if (extra !== undefined) console.log("QTP Signature: " + text, extra);
  else console.log("QTP Signature: " + text);
}

function note(item, text) {
  if (!DEBUG || !item) return;
  try {
    item.notificationMessages.replaceAsync("qtpDebug", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: ("QTP Signature: " + text).substring(0, 150),
      icon: "Icon.16x16",
      persistent: false
    });
  } catch (e) {}
}

// ---------------------------------------------------------------------------
// Sign-in and Graph lookup
// ---------------------------------------------------------------------------
var pcaPromise = null;

function getPca() {
  if (!pcaPromise) {
    pcaPromise = createNestablePublicClientApplication({
      auth: {
        clientId: CONFIG.clientId,
        authority: "https://login.microsoftonline.com/" + CONFIG.tenantId
      }
    }).catch(function (e) {
      pcaPromise = null; // allow a retry next time
      throw e;
    });
  }
  return pcaPromise;
}

async function getAccessToken(email) {
  var pca = await getPca();
  var request = { scopes: ["User.Read"] };

  var account = pca.getActiveAccount() || (pca.getAllAccounts() || [])[0];
  if (account) request.account = account;

  try {
    var result = await pca.acquireTokenSilent(request);
    return result.accessToken;
  } catch (e) {
    log("acquireTokenSilent failed, trying ssoSilent", e);
    var sso = await pca.ssoSilent({ scopes: ["User.Read"], loginHint: email });
    if (sso.account) pca.setActiveAccount(sso.account);
    return sso.accessToken;
  }
}

async function getProfile(email) {
  var token = await getAccessToken(email);
  var response = await fetch(GRAPH_URL, {
    headers: { Authorization: "Bearer " + token }
  });
  if (!response.ok) throw new Error("Graph returned " + response.status);
  return response.json();
}

// ---------------------------------------------------------------------------
// Signature template (same design as the transport rule)
// ---------------------------------------------------------------------------
function esc(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function clean(value) {
  return String(value || "").trim();
}

function row(content, style) {
  return '<tr><td style="font-size:12pt;' + (style || "") + '">' + content + "</td></tr>";
}

function buildSignature(p, email) {
  var spacer =
    '<table cellpadding="0" cellspacing="0" border="0"><tr><td style="height:20px; line-height:20px; font-size:1px;">&nbsp;</td></tr></table>';
  var rows = "";

  rows += row(esc(p.firstName) + " " + esc(p.lastName), " font-weight:bold; color:#1f3f95;");
  if (p.title) rows += row(esc(p.title), " font-weight:bold; color:#f4a11a;");
  rows += row("Quality Tractor Parts", " color:#1f3f95; font-weight:bold;");
  rows += row(
    "Units 5-7, Zone C, Mullingar Business Park, Mullingar, Co. Westmeath, Ireland. N91 ND3T",
    " color:#555555;"
  );
  rows += row(
    '<strong>Email:</strong> <a href="mailto:' + esc(email) + '" style="color:#333333; text-decoration:none;">' +
      esc(email) + "</a>"
  );

  var phoneParts = [];
  if (p.ext) phoneParts.push("<strong>Ext.:</strong> " + esc(p.ext));
  if (p.phone) phoneParts.push("<strong>Tel:</strong> " + esc(p.phone));
  if (phoneParts.length) rows += row(phoneParts.join(" | "));
  if (p.mobile) rows += row("<strong>Mob:</strong> " + esc(p.mobile));

  rows += row(
    '<a href="https://www.qtponline.com/" style="color:#1f3f95; text-decoration:none; font-weight:bold;">qtponline.com</a>'
  );

  return (
    spacer +
    '<div style="font-family:Segoe UI, Arial, sans-serif; font-size:12pt; color:#333333;">Best Regards,</div>' +
    spacer +
    '<table cellpadding="0" cellspacing="0" border="0" style="font-family:Segoe UI, Arial, sans-serif; font-size:13px; color:#333333;"><tr>' +
    '<td style="padding-right:18px; vertical-align:middle;">' +
    '<a href="https://www.qtponline.com/"><img src="' + esc(CONFIG.logoUrl) +
    '" alt="Quality Tractor Parts" width="110" style="display:block; border:0;"></a>' +
    "</td>" +
    '<td style="border-left:3px solid #1f3f95; padding-left:18px; vertical-align:middle;">' +
    '<table cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">' +
    rows + "</table>" +
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
// Office helpers (promise wrappers)
// ---------------------------------------------------------------------------
function call(fn) {
  return new Promise(function (resolve) {
    try {
      fn(function (r) { resolve(r); });
    } catch (e) {
      resolve({ status: "failed", error: e });
    }
  });
}

// ---------------------------------------------------------------------------
// Event handler - new message, reply, reply all, forward
// ---------------------------------------------------------------------------
async function onNewMessageComposeHandler(event) {
  log("handler called");
  var item;
  try {
    var mailbox = Office.context.mailbox;
    item = mailbox.item;
    var email = clean(mailbox.userProfile.emailAddress).toLowerCase();

    if (!CONFIG.clientId || !CONFIG.tenantId || CONFIG.clientId.indexOf("PASTE") === 0) {
      throw new Error("Client ID / Tenant ID not set");
    }
    if (email.slice(-ALLOWED_DOMAIN.length) !== ALLOWED_DOMAIN) {
      note(item, "not a QTP mailbox, skipped");
      event.completed();
      return;
    }

    var g = await getProfile(email);
    log("profile loaded", g);

    var person = {
      firstName: clean(g.givenName),
      lastName: clean(g.surname),
      title: clean(g.jobTitle),
      ext: clean(g.faxNumber),
      phone: clean((g.businessPhones || [])[0]),
      mobile: clean(g.mobilePhone)
    };

    if (!person.firstName && !person.lastName) {
      var parts = clean(g.displayName).split(" ");
      person.firstName = parts.shift() || "";
      person.lastName = parts.join(" ");
    }
    if (!person.firstName && !person.lastName) {
      throw new Error("no name in Entra profile");
    }

    var html = buildSignature(person, email);

    if (item.disableClientSignatureAsync) {
      await call(function (cb) { item.disableClientSignatureAsync(cb); });
    }

    if (item.internetHeaders && item.internetHeaders.setAsync) {
      var headers = {};
      headers[SIGNATURE_HEADER] = "Applied";
      await call(function (cb) { item.internetHeaders.setAsync(headers, cb); });
    }

    var r = await call(function (cb) {
      item.body.setSignatureAsync(html, { coercionType: Office.CoercionType.Html }, cb);
    });

    if (r.status === Office.AsyncResultStatus.Failed) {
      log("signature failed", r.error);
      note(item, "signature failed - " + (r.error && r.error.message));
    } else {
      note(item, "signature added for " + email);
    }
  } catch (e) {
    // Any failure: add nothing. The transport rule will add the signature instead.
    log("error", e);
    note(item, "error - " + (e && e.message));
  }
  event.completed();
}

// ---------------------------------------------------------------------------
// Register with Outlook
// ---------------------------------------------------------------------------
function registerHandler(source) {
  try {
    Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
    log("handler registered (" + source + ")");
  } catch (e) {
    log("could not register handler", e);
  }
}

registerHandler("immediate");

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(function () {
    registerHandler("onReady");
  });
}
