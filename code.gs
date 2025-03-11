function processEmailsToSheets() {
  // Label or query to filter relevant emails
  const query = "has:attachment filename:csv -label:Processed";
  
  // Get threads matching the query
  const threads = GmailApp.search(query);
  const label = GmailApp.createLabel("Processed");
  
  threads.forEach(thread => {
    const messages = thread.getMessages();
    
    messages.forEach(message => {
      const attachments = message.getAttachments();
      // subject of the messange
      const subject=message.getSubject();
      

      
      attachments.forEach(attachment => {
        if (attachment.getContentType() === "text/csv") {
          // Parse CSV content
          const csvData = Utilities.parseCsv(attachment.getDataAsString());
          
          // Create a new Google Sheet
          const sheet = SpreadsheetApp.create(subject);
          const sheetTab = sheet.getActiveSheet();
          
          // Write CSV data to the Sheet
          sheetTab.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
          
          Logger.log(`Processed: ${attachment.getName()}`);
        }
      });
    });
    
    // Label the thread to avoid reprocessing
    thread.addLabel(label);
  });
}
