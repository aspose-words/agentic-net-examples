using System;
using Aspose.Words;

class InsertDocumentWithNodeImporter
{
    static void Main()
    {
        // Path to the source DOCX file that will be inserted.
        string sourcePath = @"C:\Docs\SourceDocument.docx";

        // Create a new blank destination document.
        Document dstDoc = new Document();

        // Load the source document from the file system.
        Document srcDoc = new Document(sourcePath);

        // Create a NodeImporter that will keep the source formatting.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);

        // Destination parent node where the imported nodes will be placed.
        // Here we insert at the end of the first section's body.
        CompositeNode destinationParent = dstDoc.FirstSection.Body;

        // Keep a reference to the last inserted node so we can insert after it.
        Node lastInsertedNode = destinationParent.LastChild;

        // Iterate through all block‑level nodes in each source section.
        foreach (Section srcSection in srcDoc.Sections)
        {
            foreach (Node srcNode in srcSection.Body)
            {
                // Skip the final empty paragraph of a section (Word adds it automatically).
                if (srcNode.NodeType == NodeType.Paragraph)
                {
                    Paragraph para = (Paragraph)srcNode;
                    if (para.IsEndOfSection && !para.HasChildNodes)
                        continue;
                }

                // Import the node (deep clone) with its children.
                Node importedNode = importer.ImportNode(srcNode, true);

                // Insert the imported node after the previously inserted node.
                destinationParent.InsertAfter(importedNode, lastInsertedNode);
                lastInsertedNode = importedNode;
            }
        }

        // Save the resulting document.
        dstDoc.Save(@"C:\Docs\ResultDocument.docx");
    }
}
