Office.onReady(function() {
    document.getElementById("sendToObsidian").onclick = function() {
        sendToObsidian();
    };
});

function sendToObsidian() {
    let item = Office.context.mailbox.item;
    
    if (!item) {
        document.getElementById("status").innerText = "Error: Unable to access event details.";
        return;
    }

    let eventTitle = item.subject;
    let eventStart = item.start ? new Date(item.start).toISOString().split("T")[0] : "Unknown";
    let obsidianURI = `obsidian://new?vault=Obsidian%20Vault&file=${encodeURIComponent(eventTitle)}`;

    // Insert the link into the Outlook event
    item.body.getAsync("text", function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            let newBody = `Open in Obsidian: [Click Here](${obsidianURI})\n\n` + result.value;
            item.body.setAsync(newBody, { coercionType: "Html" }, function(setResult) {
                if (setResult.status === Office.AsyncResultStatus.Succeeded) {
                    document.getElementById("status").innerText = "✅ Link added to event!";
                } else {
                    document.getElementById("status").innerText = "❌ Failed to insert link.";
                }
            });
        } else {
            document.getElementById("status").innerText = "❌ Failed to read event body.";
        }
    });

    // Open Obsidian (optional)
    window.location.href = obsidianURI;
}
