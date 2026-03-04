using Aspose.Words;
using System;
using System.IO;
using System.Linq;

class Program
{
    static void Main()
    {
        // Load (or create) the destination document.
        Document dstDoc = new Document(); // creates a blank document.
        DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);
        dstBuilder.Writeln("Destination start"); // initial content.

        // Load the source document from a DOCX file.
        string srcPath = "Source.docx"; // <-- replace with actual path.
        Document srcDoc = new Document(srcPath);

        // Choose the node after which the source content will be inserted.
        // Here we use the first paragraph of the destination document.
        Node insertionNode = dstDoc.FirstSection.Body.FirstParagraph;

        // Verify that the insertion node is a paragraph or a table (required by the importer logic).
        if (insertionNode.NodeType != NodeType.Paragraph && insertionNode.NodeType != NodeType.Table)
            throw new ArgumentException("Insertion node must be a paragraph or a table.");

        // The parent node where new nodes will be added.
        CompositeNode parent = insertionNode.ParentNode;

        // Create a NodeImporter to efficiently import nodes from srcDoc to dstDoc.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);

        // Loop through all block‑level nodes in each section of the source document.
        foreach (Section srcSection in srcDoc.Sections.OfType<Section>())
        {
            foreach (Node srcNode in srcSection.Body)
            {
                // Skip the final empty paragraph that terminates a section.
                if (srcNode.NodeType == NodeType.Paragraph)
                {
                    Paragraph para = (Paragraph)srcNode;
                    if (para.IsEndOfSection && !para.HasChildNodes)
                        continue;
                }

                // Import the node (deep clone) into the destination document.
                Node importedNode = importer.ImportNode(srcNode, true);

                // Insert the imported node after the current insertion point.
                parent.InsertAfter(importedNode, insertionNode);
                insertionNode = importedNode; // move the insertion point forward.
            }
        }

        // Save the combined document.
        dstDoc.Save("Result.docx"); // <-- adjust output path as needed.
    }
}
