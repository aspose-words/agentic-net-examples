using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source document that will be linked.
        string sourceDocPath = @"C:\Docs\SourceDocument.docx";

        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to add content and a LINK field.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Linked content (RTF format):");

        // Insert a LINK field and obtain the strongly‑typed FieldLink object.
        FieldLink linkField = (FieldLink)builder.InsertField(FieldType.FieldLink, true);

        // Set the field to insert the linked object as Rich Text Format.
        linkField.InsertAsRtf = true;

        // Configure the field to point to the source document.
        linkField.ProgId = "Word.Document.8";
        linkField.SourceFullName = sourceDocPath;
        linkField.SourceItem = null; // No specific item; whole document.

        // Optionally enable automatic updates when the source changes.
        linkField.AutoUpdate = true;

        // Add a line break after the field for readability.
        builder.Writeln();

        // Update all fields in the document so the LINK field shows its result.
        doc.UpdateFields();

        // Prepare RTF save options – we will save the document as RTF.
        RtfSaveOptions saveOptions = new RtfSaveOptions
        {
            // Ensure the format is RTF (default, but explicit for clarity).
            SaveFormat = SaveFormat.Rtf,

            // Example: reduce size by omitting old‑reader keywords.
            ExportImagesForOldReaders = false
        };

        // Save the document to an RTF file.
        string outputPath = @"C:\Docs\ResultDocument.rtf";
        doc.Save(outputPath, saveOptions);

        // Demonstrate loading the saved RTF with RtfLoadOptions.
        RtfLoadOptions loadOptions = new RtfLoadOptions
        {
            // Recognize UTF‑8 characters if present.
            RecognizeUtf8Text = true
        };

        Document loadedDoc = new Document(outputPath, loadOptions);

        // Verify that the loaded document still contains the LINK field.
        foreach (Field field in loadedDoc.Range.Fields)
        {
            if (field.Type == FieldType.FieldLink)
            {
                FieldLink loadedLink = (FieldLink)field;
                Console.WriteLine($"Loaded LINK field InsertAsRtf = {loadedLink.InsertAsRtf}");
            }
        }
    }
}
