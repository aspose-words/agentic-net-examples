using System;
using Aspose.Words;
using Aspose.Words.Markup;

class ContentControlSplitter
{
    static void Main()
    {
        // Load the source DOCX document.
        Document sourceDoc = new Document("InputDocument.docx");

        // Get all content controls (StructuredDocumentTag nodes) in the document.
        NodeCollection contentControls = sourceDoc.GetChildNodes(NodeType.StructuredDocumentTag, true);

        int partNumber = 1;

        // Iterate through each content control and create a separate document for its contents.
        foreach (StructuredDocumentTag sdt in contentControls)
        {
            // Create a new blank document.
            Document partDoc = new Document();

            // Ensure the new document has a section and a body.
            Section section = new Section(partDoc);
            partDoc.AppendChild(section);
            Body body = new Body(partDoc);
            section.AppendChild(body);

            // Import the nodes inside the content control into the new document.
            NodeImporter importer = new NodeImporter(sourceDoc, partDoc, ImportFormatMode.KeepSourceFormatting);
            foreach (Node child in sdt.GetChildNodes(NodeType.Any, true))
            {
                Node importedNode = importer.ImportNode(child, true);
                body.AppendChild(importedNode);
            }

            // Save the split part as a separate DOCX file.
            string outputFileName = $"SplitPart_{partNumber}.docx";
            partDoc.Save(outputFileName);
            partNumber++;
        }
    }
}
