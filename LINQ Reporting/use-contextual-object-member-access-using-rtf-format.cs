using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a LINK field that will pull content from another document.
        // The field is inserted with a placeholder code; we will configure it afterwards.
        FieldLink linkField = (FieldLink)builder.InsertField(FieldType.FieldLink, true);

        // Configure the LINK field to insert the linked content as Rich Text Format (RTF).
        linkField.InsertAsRtf = true;               // Contextual member access to set RTF insertion.
        linkField.ProgId = "Word.Document.8";        // Programmatic identifier for Word documents.
        linkField.SourceFullName = @"C:\Temp\SourceDocument.docx"; // Path to the source file.
        linkField.SourceItem = null;                // No specific item within the source.
        linkField.AutoUpdate = true;                // Keep the field up‑to‑date automatically.

        // Add a line break after the field for readability.
        builder.Writeln();

        // Update all fields in the document so the LINK field resolves.
        doc.UpdateFields();

        // Prepare RTF save options.
        RtfSaveOptions saveOptions = new RtfSaveOptions
        {
            // Ensure compatibility with older readers if needed.
            ExportImagesForOldReaders = true
        };

        // Save the document as an RTF file using the specified options.
        doc.Save(@"C:\Temp\ResultDocument.rtf", saveOptions);
    }
}
