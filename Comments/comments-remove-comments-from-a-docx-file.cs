// Load the DOCX file.
var doc = new Aspose.Words.Document("input.docx");

// Instruct Aspose.Words to strip personal information, which includes comments.
doc.RemovePersonalInformation = true;

// Save the cleaned document.
doc.Save("output.docx");
