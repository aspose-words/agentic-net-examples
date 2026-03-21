using System;
using Aspose.Words;

class InsertTocExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Ensure the cursor is at the very beginning of the document.
        builder.MoveToDocumentStart();

        // Insert a Table of Contents field with default switches.
        // The switches string cannot be empty; use the standard default switches.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Update all fields so the TOC will display correct entries when opened in Word.
        doc.UpdateFields();

        // Save the document to disk.
        doc.Save("DocumentWithToc.docx");
    }
}
