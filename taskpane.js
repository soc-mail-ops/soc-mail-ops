(function () {
  "use strict";

  var DEFAULT_CONFIG = {
    prefix: "Report Phishing",
    recipients: ["iiahmad435@gmail.com","202112720@std-zuj.edu.jo"
    ],
    body: "Thank you for reporting this message.",
    non_phish_msg: "This message was identified as a simulation and was not submitted.",
    phish_msg: "Thank you. The suspicious message was reported and moved to Deleted Items.",
    env: "prod"
  };

  var state = {
    config: null,
    isProcessing: false,
    isCompleted: false
  };

  var EWS_ENVELOPE_START = "" +
    "<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" " +
    "xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\" " +
    "xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" " +
    "xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">" +
    "<soap:Header><t:RequestServerVersion Version=\"Exchange2013\"/></soap:Header><soap:Body>";

  var EWS_ENVELOPE_END = "</soap:Body></soap:Envelope>";

  function bootstrap() {
    bindReportButton();

    if (typeof Office === "undefined" || !Office.onReady) {
      setStatus("Office.js is unavailable. Open this page inside Outlook.", "error");
      disableReportButton(true);
      return;
    }

    Office.onReady(function (info) {
      if (!info || info.host !== Office.HostType.Outlook) {
        setStatus("This add-in can only run inside Outlook.", "error");
        disableReportButton(true);
        return;
      }

      state.config = parseRuntimeConfig();
      applyConfigToUi(state.config);
      setStepText(1, "Runtime configuration loaded");
      runReportWorkflow();
    });
  }

  function bindReportButton() {
    var button = getEl("reportButton");
    if (!button) {
      return;
    }

    button.onclick = function () {
      if (!state.isProcessing && !state.isCompleted) {
        runReportWorkflow();
      }
    };
  }

  function runReportWorkflow() {
    if (state.isProcessing || state.isCompleted) {
      return;
    }

    state.isProcessing = true;
    disableReportButton(true);
    setResult("");
    setStatus("Inspecting message headers...", "pending");
    setStepText(2, "Inspecting X-PHISHTEST internet header");

    var itemId = getCurrentEwsItemId();
    if (!itemId) {
      failWorkflow("Unable to read the selected Outlook message.");
      return;
    }

    fetchSimulationHeader(itemId, function (isSimulation) {
      if (isSimulation) {
        setStepText(2, "Simulation header detected");
        setStepText(3, "Security submission skipped for simulation");
        setStepText(4, "Original message kept in mailbox");
        finishWorkflow(state.config.non_phish_msg, true);
        return;
      }

      setStepText(2, "No simulation header found");
      setStatus("Extracting MIME content...", "pending");
      fetchMimeContent(itemId, function (mimeContentBase64) {
        setStepText(3, "Sending phishing report to security recipients");
        setStatus("Submitting report...", "pending");

        sendPhishingReport(mimeContentBase64, function () {
          setStepText(3, "Security report sent");
          setStepText(4, "Moving original message to Deleted Items");
          setStatus("Moving message to Deleted Items...", "pending");

          moveMessageToDeletedItems(itemId, function () {
            setStepText(4, "Original message moved to Deleted Items");
            finishWorkflow(state.config.phish_msg, false);
          }, function (errorMessage) {
            failWorkflow("The report was sent, but the original message could not be moved.", errorMessage);
          });
        }, function (errorMessage) {
          failWorkflow("Failed to submit the phishing report.", errorMessage);
        });
      }, function (errorMessage) {
        failWorkflow("Failed to extract the MIME content from the message.", errorMessage);
      });
    }, function (errorMessage) {
      failWorkflow("Failed to read the message headers.", errorMessage);
    });
  }

  function finishWorkflow(userMessage, isSimulation) {
    state.isProcessing = false;
    state.isCompleted = true;

    if (isSimulation) {
      setStatus("Simulation message detected.", "success");
    } else {
      setStatus("Phishing report completed.", "success");
    }

    setResult(userMessage || "Done.");
    disableReportButton(true);
  }

  function failWorkflow(summary, details) {
    state.isProcessing = false;
    disableReportButton(false);
    setStatus(summary, "error");
    setResult(summary);

    if (details) {
      setResult(summary + " " + details);
      logToConsole("Workflow error details: " + details);
    }
  }

  function fetchSimulationHeader(itemId, onSuccess, onError) {
    var body = "<m:GetItem><m:ItemShape><t:BaseShape>IdOnly</t:BaseShape><t:AdditionalProperties><t:ExtendedFieldURI DistinguishedPropertySetId=\"InternetHeaders\" PropertyName=\"X-PHISHTEST\" PropertyType=\"String\" /></t:AdditionalProperties></m:ItemShape><m:ItemIds><t:ItemId Id=\"" + xmlEscape(itemId) + "\" /></m:ItemIds></m:GetItem>";

    makeEwsRequest(wrapInExchange2013Envelope(body), function (responseXml) {
      onSuccess(hasXPhishTestHeader(responseXml));
    }, onError);
  }

  function fetchMimeContent(itemId, onSuccess, onError) {
    var body = "" +
      "<m:GetItem>" +
      "<m:ItemShape>" +
      "<t:BaseShape>IdOnly</t:BaseShape>" +
      "<t:IncludeMimeContent>true</t:IncludeMimeContent>" +
      "</m:ItemShape>" +
      "<m:ItemIds><t:ItemId Id=\"" + xmlEscape(itemId) + "\" /></m:ItemIds>" +
      "</m:GetItem>";

    makeEwsRequest(wrapInExchange2013Envelope(body), function (responseXml) {
      var mimeContent = extractMimeContent(responseXml);
      if (!mimeContent) {
        onError("The MIME payload was empty.");
        return;
      }

      onSuccess(mimeContent);
    }, onError);
  }

  function sendPhishingReport(mimeContentBase64, onSuccess, onError) {
    var recipients = state.config.recipients || [];
    if (!recipients.length) {
      onError("No recipients were provided in the payload.");
      return;
    }

    var subject = "[" + state.config.prefix + "] " + getCurrentItemSubject();
    var body = buildCreateItemBody(subject, state.config.body, recipients, mimeContentBase64);
    makeEwsRequest(wrapInExchange2013Envelope(body), function () {
      onSuccess();
    }, onError);
  }

  function moveMessageToDeletedItems(itemId, onSuccess, onError) {
    var body = "<MoveItem xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\"><ToFolderId><t:DistinguishedFolderId Id=\"deleteditems\"/></ToFolderId><ItemIds><t:ItemId Id=\"" + xmlEscape(itemId) + "\"/></ItemIds></MoveItem>";

    makeEwsRequest(wrapInExchange2013Envelope(body), function () {
      onSuccess();
    }, onError);
  }

  function makeEwsRequest(soapEnvelope, onSuccess, onError) {
    if (!Office || !Office.context || !Office.context.mailbox || !Office.context.mailbox.makeEwsRequestAsync) {
      onError("makeEwsRequestAsync is not available in this client.");
      return;
    }

    Office.context.mailbox.makeEwsRequestAsync(soapEnvelope, function (asyncResult) {
      if (!asyncResult || asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        onError(formatOfficeError(asyncResult));
        return;
      }

      var xmlText = asyncResult.value || "";
      var faultMessage = extractSoapFaultMessage(xmlText);
      if (faultMessage) {
        onError(faultMessage);
        return;
      }

      var responseFailure = extractResponseFailure(xmlText);
      if (responseFailure) {
        onError(responseFailure);
        return;
      }

      onSuccess(xmlText);
    });
  }

  function buildCreateItemBody(subject, reportBody, recipients, mimeContentBase64) {
    var recipientsXml = "";
    var i;

    for (i = 0; i < recipients.length; i += 1) {
      recipientsXml += "<t:Mailbox><t:EmailAddress>" + xmlEscape(recipients[i]) + "</t:EmailAddress></t:Mailbox>";
    }

    var cleanedMime = String(mimeContentBase64 || "").replace(/\s+/g, "");

    return "" +
      "<m:CreateItem MessageDisposition=\"SendOnly\">" +
      "<m:Items>" +
      "<t:Message>" +
      "<t:Subject>" + xmlEscape(subject) + "</t:Subject>" +
      "<t:Body BodyType=\"Text\">" + xmlEscape(reportBody) + "</t:Body>" +
      "<t:ToRecipients>" + recipientsXml + "</t:ToRecipients>" +
      "<t:Attachments>" +
      "<t:FileAttachment>" +
      "<t:Name>reported-message.eml</t:Name>" +
      "<t:ContentType>message/rfc822</t:ContentType>" +
      "<t:IsInline>false</t:IsInline>" +
      "<t:Content>" + cleanedMime + "</t:Content>" +
      "</t:FileAttachment>" +
      "</t:Attachments>" +
      "</t:Message>" +
      "</m:Items>" +
      "</m:CreateItem>";
  }

  function wrapInExchange2013Envelope(innerBodyXml) {
    return EWS_ENVELOPE_START + innerBodyXml + EWS_ENVELOPE_END;
  }

  function hasXPhishTestHeader(responseXml) {
    var documentXml = parseXml(responseXml);
    if (!documentXml) {
      return /X-PHISHTEST/i.test(responseXml || "");
    }

    var extendedProperties = getElementsByLocalName(documentXml, "ExtendedProperty");
    var i;
    for (i = 0; i < extendedProperties.length; i += 1) {
      var fieldUris = getElementsByLocalName(extendedProperties[i], "ExtendedFieldURI");
      if (!fieldUris || !fieldUris.length) {
        continue;
      }

      var propertyName = fieldUris[0].getAttribute("PropertyName") || fieldUris[0].getAttribute("propertyname");
      if (propertyName && propertyName.toUpperCase() === "X-PHISHTEST") {
        return true;
      }
    }

    return false;
  }

  function extractMimeContent(responseXml) {
    var documentXml = parseXml(responseXml);
    if (!documentXml) {
      return "";
    }

    var mimeNodes = getElementsByLocalName(documentXml, "MimeContent");
    if (!mimeNodes || !mimeNodes.length) {
      return "";
    }

    return trim(getNodeText(mimeNodes[0])).replace(/\s+/g, "");
  }

  function extractSoapFaultMessage(responseXml) {
    var documentXml = parseXml(responseXml);
    if (!documentXml) {
      return "";
    }

    var faultStrings = getElementsByLocalName(documentXml, "faultstring");
    if (faultStrings && faultStrings.length) {
      return trim(getNodeText(faultStrings[0]));
    }

    var faults = getElementsByLocalName(documentXml, "Fault");
    if (faults && faults.length) {
      return trim(getNodeText(faults[0]));
    }

    return "";
  }

  function extractResponseFailure(responseXml) {
    var documentXml = parseXml(responseXml);
    if (!documentXml) {
      return "";
    }

    var responseCodes = getElementsByLocalName(documentXml, "ResponseCode");
    var i;
    for (i = 0; i < responseCodes.length; i += 1) {
      var code = trim(getNodeText(responseCodes[i]));
      if (code && code !== "NoError") {
        var messageTexts = getElementsByLocalName(documentXml, "MessageText");
        var message = messageTexts && messageTexts.length ? trim(getNodeText(messageTexts[0])) : "";
        if (message) {
          return code + ": " + message;
        }

        return code;
      }
    }

    return "";
  }

  function parseRuntimeConfig() {
    var candidates = collectPayloadCandidates();
    var i;
    var parsed;

    for (i = 0; i < candidates.length; i += 1) {
      parsed = decodePayloadCandidate(candidates[i]);
      if (parsed) {
        return normalizeConfig(parsed);
      }
    }

    return normalizeConfig(DEFAULT_CONFIG);
  }

  function collectPayloadCandidates() {
    var candidates = [];

    addCandidate(candidates, window.location.hash);

    var href = window.location.href || "";
    var hashIndex = href.indexOf("#");
    if (hashIndex > -1 && hashIndex < href.length - 1) {
      addCandidate(candidates, href.substring(hashIndex + 1));
    }

    var searchParams = parseQueryString((window.location.search || "").replace(/^\?/, ""));
    addCandidate(candidates, searchParams.et);
    addCandidate(candidates, searchParams.payload);
    addCandidate(candidates, searchParams.config);

    var key;
    for (key in searchParams) {
      if (!searchParams.hasOwnProperty(key)) {
        continue;
      }

      if (searchParams[key] && searchParams[key].charAt(0) === "#") {
        addCandidate(candidates, searchParams[key].substring(1));
      }
    }

    var hashParams = parseQueryString((window.location.hash || "").replace(/^#/, ""));
    addCandidate(candidates, hashParams.et);
    addCandidate(candidates, hashParams.payload);
    addCandidate(candidates, hashParams.config);

    return candidates;
  }

  function addCandidate(candidates, value) {
    if (typeof value !== "string") {
      return;
    }

    var normalized = sanitizeCandidate(value);
    if (!normalized) {
      return;
    }

    var i;
    for (i = 0; i < candidates.length; i += 1) {
      if (candidates[i] === normalized) {
        return;
      }
    }

    candidates.push(normalized);
  }

  function sanitizeCandidate(value) {
    var normalized = trim(String(value || ""));
    if (!normalized) {
      return "";
    }

    normalized = normalized.replace(/^[#?]+/, "");
    normalized = safeDecodeURIComponent(normalized);
    normalized = normalized.replace(/^[#?]+/, "");

    if (normalized.indexOf("et=") === 0 || normalized.indexOf("payload=") === 0 || normalized.indexOf("config=") === 0) {
      normalized = normalized.substring(normalized.indexOf("=") + 1);
    }

    normalized = normalized.replace(/^[#?]+/, "");

    if (normalized.indexOf("#") > -1) {
      normalized = normalized.substring(normalized.lastIndexOf("#") + 1);
    }

    if (normalized.indexOf("&") > -1) {
      normalized = normalized.split("&")[0];
    }

    return trim(normalized);
  }

  function decodePayloadCandidate(candidate) {
    if (!candidate) {
      return null;
    }

    var directText = safeDecodeURIComponent(candidate);
    if (directText && directText.charAt(0) === "{") {
      try {
        return JSON.parse(directText);
      } catch (jsonParseError) {
        return null;
      }
    }

    var decodedJsonText = decodeBase64Utf8(directText);
    if (!decodedJsonText) {
      return null;
    }

    try {
      return JSON.parse(decodedJsonText);
    } catch (error) {
      return null;
    }
  }

  function decodeBase64Utf8(base64Value) {
    var normalized = String(base64Value || "").replace(/\s+/g, "").replace(/-/g, "+").replace(/_/g, "/");
    if (!normalized) {
      return "";
    }

    while (normalized.length % 4 !== 0) {
      normalized += "=";
    }

    var binaryString;
    try {
      binaryString = window.atob(normalized);
    } catch (decodeError) {
      return "";
    }

    try {
      if (window.TextDecoder && window.Uint8Array) {
        var bytes = new Uint8Array(binaryString.length);
        var i;
        for (i = 0; i < binaryString.length; i += 1) {
          bytes[i] = binaryString.charCodeAt(i);
        }

        return new TextDecoder("utf-8").decode(bytes);
      }
    } catch (textDecoderError) {
      logToConsole("TextDecoder fallback used.");
    }

    var encoded = "";
    var j;
    for (j = 0; j < binaryString.length; j += 1) {
      var code = binaryString.charCodeAt(j).toString(16);
      if (code.length < 2) {
        code = "0" + code;
      }

      encoded += "%" + code;
    }

    try {
      return decodeURIComponent(encoded);
    } catch (fallbackError) {
      return binaryString;
    }
  }

  function normalizeConfig(source) {
    var merged = {
      prefix: DEFAULT_CONFIG.prefix,
      recipients: DEFAULT_CONFIG.recipients.slice(0),
      body: DEFAULT_CONFIG.body,
      non_phish_msg: DEFAULT_CONFIG.non_phish_msg,
      phish_msg: DEFAULT_CONFIG.phish_msg,
      env: DEFAULT_CONFIG.env
    };

    if (!source || typeof source !== "object") {
      return merged;
    }

    if (typeof source.prefix === "string" && trim(source.prefix)) {
      merged.prefix = trim(source.prefix);
    }

    if (typeof source.body === "string" && trim(source.body)) {
      merged.body = trim(source.body);
    }

    if (typeof source.non_phish_msg === "string" && trim(source.non_phish_msg)) {
      merged.non_phish_msg = trim(source.non_phish_msg);
    }

    if (typeof source.phish_msg === "string" && trim(source.phish_msg)) {
      merged.phish_msg = trim(source.phish_msg);
    }

    if (typeof source.env === "string" && trim(source.env)) {
      merged.env = trim(source.env);
    }

    if (source.recipients && source.recipients.length) {
      merged.recipients = [];

      var i;
      for (i = 0; i < source.recipients.length; i += 1) {
        if (typeof source.recipients[i] === "string" && trim(source.recipients[i])) {
          merged.recipients.push(trim(source.recipients[i]));
        }
      }
    }

    return merged;
  }

  function applyConfigToUi(config) {
    var normalizedConfig = config || DEFAULT_CONFIG;

    setText("titleText", normalizedConfig.prefix);
    setText("reportButton", normalizedConfig.prefix);
    setText("subtitleText", normalizedConfig.body);
    setText("resultText", "");

    var envBadge = getEl("envBadge");
    if (envBadge) {
      var envName = trim(normalizedConfig.env || "prod").toLowerCase();
      if (envName && envName !== "prod") {
        envBadge.hidden = false;
        envBadge.textContent = envName;
      } else {
        envBadge.hidden = true;
      }
    }

    if (isRtlText(normalizedConfig.prefix + " " + normalizedConfig.body + " " + normalizedConfig.phish_msg + " " + normalizedConfig.non_phish_msg)) {
      document.documentElement.setAttribute("dir", "rtl");
      document.documentElement.setAttribute("lang", "ar");
    } else {
      document.documentElement.setAttribute("dir", "ltr");
      document.documentElement.setAttribute("lang", "en");
    }
  }

  function getCurrentEwsItemId() {
    var mailbox = Office.context && Office.context.mailbox;
    var item = mailbox && mailbox.item;

    if (!item || !item.itemId) {
      return "";
    }

    var itemId = item.itemId;

    if (mailbox.convertToEwsId && Office.MailboxEnums && Office.MailboxEnums.RestVersion && Office.MailboxEnums.RestVersion.v2_0) {
      try {
        itemId = mailbox.convertToEwsId(item.itemId, Office.MailboxEnums.RestVersion.v2_0);
      } catch (conversionError) {
        itemId = item.itemId;
      }
    }

    return itemId;
  }

  function getCurrentItemSubject() {
    var mailbox = Office.context && Office.context.mailbox;
    var item = mailbox && mailbox.item;
    var subject = item && item.subject ? item.subject : "(No subject)";
    return trim(subject);
  }

  function parseQueryString(queryString) {
    var result = {};
    if (!queryString) {
      return result;
    }

    var pairs = queryString.split("&");
    var i;
    for (i = 0; i < pairs.length; i += 1) {
      if (!pairs[i]) {
        continue;
      }

      var separator = pairs[i].indexOf("=");
      var rawKey = separator > -1 ? pairs[i].substring(0, separator) : pairs[i];
      var rawValue = separator > -1 ? pairs[i].substring(separator + 1) : "";

      var key = safeDecodeURIComponent(rawKey).toLowerCase();
      var value = safeDecodeURIComponent(rawValue);

      if (key) {
        result[key] = value;
      }
    }

    return result;
  }

  function parseXml(xmlText) {
    if (!xmlText) {
      return null;
    }

    try {
      if (window.DOMParser) {
        return (new DOMParser()).parseFromString(xmlText, "text/xml");
      }

      if (window.ActiveXObject) {
        var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
        xmlDoc.async = false;
        xmlDoc.loadXML(xmlText);
        return xmlDoc;
      }
    } catch (error) {
      return null;
    }

    return null;
  }

  function getElementsByLocalName(xmlNode, localName) {
    if (!xmlNode || !localName) {
      return [];
    }

    if (xmlNode.getElementsByTagNameNS) {
      var byNamespace = xmlNode.getElementsByTagNameNS("*", localName);
      if (byNamespace && byNamespace.length) {
        return byNamespace;
      }
    }

    var prefixedType = xmlNode.getElementsByTagName("t:" + localName);
    if (prefixedType && prefixedType.length) {
      return prefixedType;
    }

    var prefixedMessage = xmlNode.getElementsByTagName("m:" + localName);
    if (prefixedMessage && prefixedMessage.length) {
      return prefixedMessage;
    }

    return xmlNode.getElementsByTagName(localName);
  }

  function getNodeText(xmlNode) {
    if (!xmlNode) {
      return "";
    }

    if (typeof xmlNode.textContent === "string") {
      return xmlNode.textContent;
    }

    if (typeof xmlNode.text === "string") {
      return xmlNode.text;
    }

    return "";
  }

  function formatOfficeError(asyncResult) {
    if (!asyncResult) {
      return "Unknown Office.js async error.";
    }

    if (asyncResult.error && asyncResult.error.message) {
      return asyncResult.error.message;
    }

    if (asyncResult.error && asyncResult.error.name) {
      return asyncResult.error.name;
    }

    return "Office.js operation failed.";
  }

  function xmlEscape(value) {
    return String(value || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  function setStatus(message, kind) {
    var status = getEl("statusText");
    if (!status) {
      return;
    }

    status.className = "status";
    if (kind === "success") {
      status.className += " status-success";
    } else if (kind === "error") {
      status.className += " status-error";
    } else {
      status.className += " status-pending";
    }

    status.textContent = message;
  }

  function setResult(message) {
    setText("resultText", message || "");
  }

  function setStepText(stepNumber, message) {
    setText("step-" + stepNumber, stepNumber + ". " + message);
  }

  function setText(elementId, text) {
    var el = getEl(elementId);
    if (!el) {
      return;
    }

    el.textContent = text;
  }

  function disableReportButton(disabled) {
    var button = getEl("reportButton");
    if (!button) {
      return;
    }

    button.disabled = !!disabled;
  }

  function safeDecodeURIComponent(value) {
    if (typeof value !== "string") {
      return "";
    }

    try {
      return decodeURIComponent(value);
    } catch (error) {
      return value;
    }
  }

  function trim(value) {
    return String(value || "").replace(/^\s+|\s+$/g, "");
  }

  function getEl(id) {
    return document.getElementById(id);
  }

  function isRtlText(value) {
    return /[\u0591-\u07FF\uFB1D-\uFDFD\uFE70-\uFEFC]/.test(value || "");
  }

  function logToConsole(message) {
    if (window.console && console.log) {
      console.log(message);
    }
  }

  bootstrap();
})();
