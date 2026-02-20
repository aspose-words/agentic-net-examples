using System;
using Aspose.Words;
using Aspose.Words.Markup;

class ContentControlDemo
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph to host the content control.
        builder.Writeln("Below is a plain‑text content control:");

        // Create a plain‑text Structured Document Tag (content control) at the inline level.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",          // Friendly name shown in the UI.
            Tag = "CustomerNameTag",         // Tag used for identification.
            PlaceholderName = "Enter name", // Placeholder text displayed when empty.
            LockContents = false,           // Allow editing of the content.
            LockContentControl = false      // Allow deletion of the control.
        };

        // Insert the content control into the document.
        builder.InsertNode(sdt);

        // Insert some default text inside the content control.
        builder.MoveTo(sdt);
        builder.Write("John Doe");

        // Add another paragraph after the content control.
        builder.Writeln();
        builder.Writeln("Another content control (checkbox):");

        // Create a checkbox content control.
        StructuredDocumentTag checkBox = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline)
        {
            Title = "AgreeTerms",
            Tag = "AgreeTermsTag",
            Checked = false,
            LockContentControl = false,
            LockContents = false
        };

        // Insert the checkbox control.
        builder.InsertNode(checkBox);

        // Save the document to disk.
        doc.Save("ContentControlDemo.docx");
    }
}
