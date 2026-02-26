using System;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("Input.docx");

        // Retrieve all content controls (StructuredDocumentTag nodes) in the document.
        NodeCollection contentControls = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);

        // Update each content control.
        foreach (StructuredDocumentTag sdt in contentControls)
        {
            // Example update: replace the control's contents with a string that includes its title.
            string updatedText = $"Updated content for \"{sdt.Title}\"";

            // Remove any existing child nodes (e.g., previous text runs).
            sdt.RemoveAllChildren();

            // Insert a new Run node containing the updated text.
            Run run = new Run(doc, updatedText);
            sdt.AppendChild(run);
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
