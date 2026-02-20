using System;
using Aspose.Words;
using Aspose.Words.Markup;

class UpdateContentControls
{
    static void Main()
    {
        // Load the existing DOCX file.
        Document doc = new Document("Input.docx");

        // Iterate through all content controls (structured document tags) in the document.
        foreach (StructuredDocumentTag sdt in doc.Range.StructuredDocumentTags)
        {
            // Update a content control identified by its Tag property.
            if (sdt.Tag == "CustomerName")
            {
                UpdateTagText(sdt, "John Doe");
            }
            // Update a content control identified by its Title property.
            else if (sdt.Title == "CurrentDate")
            {
                UpdateTagText(sdt, DateTime.Today.ToString("MMMM dd, yyyy"));
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }

    /// <summary>
    /// Replaces the contents of a StructuredDocumentTag with the specified text.
    /// </summary>
    private static void UpdateTagText(StructuredDocumentTag tag, string newText)
    {
        // Remove any existing child nodes (the current content of the control).
        tag.RemoveAllChildren();

        // Create a new Run node that holds the replacement text.
        Run run = new Run(tag.Document, newText);

        // Append the Run to the content control.
        tag.AppendChild(run);
    }
}
