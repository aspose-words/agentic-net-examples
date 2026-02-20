using System;
using Aspose.Words;
using Aspose.Words.Markup;

class ReplaceInStructuredDocumentTags
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Text to find and its replacement.
        string oldText = "PLACEHOLDER";
        string newText = "Actual Value";

        // Get the collection of structured document tags (SDTs) in the document.
        StructuredDocumentTagCollection sdtCollection = doc.Range.StructuredDocumentTags;

        // Iterate through each SDT. Use the concrete StructuredDocumentTag class – it exposes the Range property.
        foreach (StructuredDocumentTag sdt in sdtCollection)
        {
            // Perform a simple string replace inside the SDT's range.
            sdt.Range.Replace(oldText, newText);
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
