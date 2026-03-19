Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("reportButton").onclick = executeSocReport;
    }
});

function executeSocReport() {
    const btn = document.getElementById("reportButton");
    const statusEl = document.getElementById("statusText");
    
    btn.disabled = true;
    statusEl.textContent = "Routing payload to SOC...";
    statusEl.className = "status";

    const itemId = Office.context.mailbox.item.itemId;
    const socEmail = "AbuZaid@nourtest.onmicrosoft.com";

    // 1. Forward the email natively via EWS
    const forwardEnvelope = `<?xml version="1.0" encoding="utf-8"?>
    <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                   xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                   xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                   xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
      <soap:Header>
        <t:RequestServerVersion Version="Exchange2013" />
      </soap:Header>
      <soap:Body>
        <m:CreateItem MessageDisposition="SendAndSaveCopy">
          <m:Items>
            <t:ForwardItem>
              <t:ToRecipients>
                <t:Mailbox><t:EmailAddress>${socEmail}</t:EmailAddress></t:Mailbox>
              </t:ToRecipients>
              <t:ReferenceItemId Id="${itemId}" />
              <t:NewBodyContent BodyType="Text">This email was submitted by a user for SOC analysis.</t:NewBodyContent>
            </t:ForwardItem>
          </m:Items>
        </m:CreateItem>
      </soap:Body>
    </soap:Envelope>`;

    Office.context.mailbox.makeEwsRequestAsync(forwardEnvelope, (forwardResult) => {
        if (forwardResult.status !== Office.AsyncResultStatus.Succeeded || forwardResult.value.includes("ResponseClass=\"Error\"")) {
            statusEl.textContent = "Error: Failed to route to SOC.";
            statusEl.className = "status error";
            btn.disabled = false;
            console.error(forwardResult);
            return;
        }

        // 2. Move original email to Deleted Items
        const moveEnvelope = `<?xml version="1.0" encoding="utf-8"?>
        <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                       xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                       xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                       xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
          <soap:Header>
            <t:RequestServerVersion Version="Exchange2013" />
          </soap:Header>
          <soap:Body>
            <m:MoveItem>
              <m:ToFolderId><t:DistinguishedFolderId Id="deleteditems"/></m:ToFolderId>
              <m:ItemIds><t:ItemId Id="${itemId}"/></m:ItemIds>
            </m:MoveItem>
          </soap:Body>
        </soap:Envelope>`;

        Office.context.mailbox.makeEwsRequestAsync(moveEnvelope, (moveResult) => {
            if (moveResult.status === Office.AsyncResultStatus.Succeeded && !moveResult.value.includes("ResponseClass=\"Error\"")) {
                statusEl.textContent = "Reported to SOC and moved to Deleted Items.";
                statusEl.className = "status success";
            } else {
                statusEl.textContent = "Reported to SOC, but failed to delete original email.";
                statusEl.className = "status error";
            }
        });
    });
}