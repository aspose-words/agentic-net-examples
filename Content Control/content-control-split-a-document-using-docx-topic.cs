using System;
using Aspose.Words;
using Aspose.Words.Markup;

class SplitByContentControl
{
    static void Main()
    {
        // Load the source DOCX document.
        Document sourceDoc = new Document("Input.docx");

        // Get all content controls (StructuredDocumentTag nodes) in the document.
        NodeCollection contentControls = sourceDoc.GetChildNodes(NodeType.StructuredDocumentTag, true);

        int partNumber = 1;

        // Iterate through each content control and save its contents as a separate document.
        foreach (StructuredDocumentTag sdt in contentControls)
        {
            // Create a new empty document that will hold the extracted part.
            Document partDoc = new Document();

            // Ensure the new document has at least one section/body.
            partDoc.EnsureMinimum();

            // Import the content control node (including its children) into the new document.
            NodeImporter importer = new NodeImporter(sourceDoc, partDoc, ImportFormatMode.KeepSourceFormatting);
            Node importedNode = importer.ImportNode(sdt, true);

            // Append the imported node to the body of the first section.
            partDoc.FirstSection.Body.AppendChild(importedNode);

            // Save the extracted part to a separate DOCX file.
            string outputPath = $"Part_{partNumber}.docx";
            partDoc.Save(outputPath);

            partNumber++;
        }
    }
}
