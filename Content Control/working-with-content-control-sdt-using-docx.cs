using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Replacing; // Added for FindReplaceOptions

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a plain‑text content control (SDT) at the current cursor position.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",          // Friendly name.
            Tag = "CustomerNameTag",         // Tag identifier.
            LockContents = true,            // Prevent editing of the content.
            LockContentControl = false      // Allow the control itself to be deleted if needed.
        };
        builder.InsertNode(sdt);

        // Add placeholder text inside the SDT.
        sdt.AppendChild(new Run(doc, "Enter name here"));

        // Save the document with the content control.
        doc.Save("ContentControlCreated.docx");

        // -----------------------------------------------------------------
        // Load the document we just saved and modify the content control.
        Document loadedDoc = new Document("ContentControlCreated.docx");

        // Retrieve the SDT by its title.
        IStructuredDocumentTag tag = loadedDoc.Range.StructuredDocumentTags.GetByTitle("CustomerName");
        if (tag != null && tag.Node is StructuredDocumentTag sdtNode)
        {
            // Replace the placeholder text with an actual name using Find/Replace on the SDT's range.
            FindReplaceOptions options = new FindReplaceOptions { MatchCase = false };
            sdtNode.Range.Replace("Enter name here", "John Doe", options);

            // Optionally lock the control to prevent deletion as well.
            sdtNode.LockContentControl = true;
        }

        // Save the modified document.
        loadedDoc.Save("ContentControlModified.docx");
    }
}
