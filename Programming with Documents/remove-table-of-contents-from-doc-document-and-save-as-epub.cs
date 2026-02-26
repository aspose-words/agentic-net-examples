using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields; // Added for Field and FieldType

class RemoveTocAndSaveAsEpub
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("input.docx");

        // Remove all Table of Contents (TOC) fields from the document.
        // Iterate backwards to avoid modifying the collection while enumerating.
        for (int i = doc.Range.Fields.Count - 1; i >= 0; i--)
        {
            Field field = doc.Range.Fields[i];
            if (field.Type == FieldType.FieldTOC)
                field.Remove();
        }

        // Prepare save options for EPUB format.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Epub)
        {
            ExportDocumentProperties = false // Explicitly set for clarity.
        };

        // Save the modified document as EPUB.
        doc.Save("output.epub", saveOptions);
    }
}
