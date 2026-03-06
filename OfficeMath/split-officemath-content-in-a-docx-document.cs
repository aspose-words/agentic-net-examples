using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

class OfficeMathSplitter
{
    static void Main()
    {
        // Load the source document that contains OfficeMath (equation) objects.
        Document srcDoc = new Document("Input.docx");

        // Get all OfficeMath nodes in the document (including those inside tables, paragraphs, etc.).
        NodeCollection mathNodes = srcDoc.GetChildNodes(NodeType.OfficeMath, true);

        // Iterate through each OfficeMath node and save it as a separate document.
        for (int i = 0; i < mathNodes.Count; i++)
        {
            // Create a new blank document that will hold a single OfficeMath object.
            Document partDoc = new Document();
            partDoc.EnsureMinimum(); // Guarantees a section, body and a paragraph exist.

            // Import the OfficeMath node from the source document into the new document.
            NodeImporter importer = new NodeImporter(srcDoc, partDoc, ImportFormatMode.KeepSourceFormatting);
            Node importedMath = importer.ImportNode(mathNodes[i], true);

            // Append the imported OfficeMath node to the first paragraph of the new document.
            Paragraph para = partDoc.FirstSection.Body.FirstParagraph;
            para.AppendChild(importedMath);

            // Build a file name for the split part, e.g., "OfficeMath_1.docx", "OfficeMath_2.docx", etc.
            string outFileName = $"OfficeMath_{i + 1}.docx";

            // Save the part document to disk.
            partDoc.Save(outFileName);
        }

        Console.WriteLine("OfficeMath splitting completed.");
    }
}
