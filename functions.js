// functions.js

Office.onReady(function (info) {
    // Check if the add-in is running in Outlook
    if (info.host === Office.HostType.Outlook) {
        // Perform initialization tasks if necessary
    }
});

// Function to be called when the add-in's button is clicked
function createObsidianNote() {
    // Get the current item (appointment)
    var item = Office.context.mailbox.item;

    // Extract relevant details from the appointment
    var subject = item.subject || 'Untitled Event';
    var start = item.start ? item.start.toLocaleString() : 'No start time';
    var end = item.end ? item.end.toLocaleString() : 'No end time';
    var location = item.location || 'No location';
    var body = item.body ? item.body.getAsync("text", function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            return result.value;
        } else {
            return 'No details';
        }
    }) : 'No details';

    // Format the note content
    var noteContent = `# ${subject}\n\n**Start:** ${start}\n**End:** ${end}\n**Location:** ${location}\n\n## Details\n${body}`;

    // Encode the note content for the Obsidian URI
    var encodedContent = encodeURIComponent(noteContent);

    // Create the Obsidian URI
    var obsidianUri = `obsidian://new?content=${encodedContent}`;

    // Open the Obsidian URI
    window.location.href = obsidianUri;
}
