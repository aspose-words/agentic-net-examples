using System;
using Aspose.Words;
using Aspose.Words.Markup;

class UpdateContentControls
{
    static void Main()
    {
        // Load an existing Word document.
        Document doc = new Document("Input.docx");

        // Iterate through all structured document tags (content controls) in the document.
        StructuredDocumentTagCollection sdtCollection = doc.Range.StructuredDocumentTags;
        foreach (StructuredDocumentTag sdt in sdtCollection)
        {
            // Remove any existing child nodes inside the content control.
            sdt.RemoveAllChildren();

            // Insert the new text as a Run node.
            Run run = new Run(doc, "Updated content");
            sdt.AppendChild(run);

            // Optionally, change the title of the content control.
            // sdt.Title = "NewTitle";

            // Optionally, change the tag (identifier) of the content control.
            // sdt.Tag = "NewTag";
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
