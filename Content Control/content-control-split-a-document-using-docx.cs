using System;
using Aspose.Words;
using Aspose.Words.Markup;

class SplitDocumentByContentControls
{
    static void Main()
    {
        // Load the source DOCX document.
        string inputPath = "input.docx";
        Document sourceDoc = new Document(inputPath);

        // Get all content controls (StructuredDocumentTag nodes) in the document.
        NodeCollection contentControls = sourceDoc.GetChildNodes(NodeType.StructuredDocumentTag, true);

        int partIndex = 1;
        foreach (StructuredDocumentTag sdt in contentControls)
        {
            // Create a new empty document for the current part.
            Document partDoc = new Document();

            // Import the content control node into the new document.
            Node importedNode = partDoc.ImportNode(sdt, true);

            // Append the imported node to the body of the new document.
            partDoc.FirstSection.Body.AppendChild(importedNode);

            // Save the part as a separate DOCX file.
            string outputPath = $"output_part_{partIndex}.docx";
            partDoc.Save(outputPath);

            partIndex++;
        }
    }
}
