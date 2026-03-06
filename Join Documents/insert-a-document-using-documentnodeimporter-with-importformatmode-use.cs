using System;
using Aspose.Words;

public class DocumentInserter
{
    // Inserts the contents of srcPath into dstPath using NodeImporter with UseDestinationStyles.
    // The result is saved to outputPath.
    public static void InsertDocumentUsingNodeImporter(string dstPath, string srcPath, string outputPath)
    {
        // Load the destination document.
        Document dstDoc = new Document(dstPath);

        // Load the source document.
        Document srcDoc = new Document(srcPath);

        // Choose an insertion point – after the first paragraph of the destination.
        Node insertionDestination = dstDoc.FirstSection.Body.FirstParagraph;

        // Verify that the insertion point is a paragraph or a table.
        if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
            throw new ArgumentException("Insertion point must be a paragraph or a table.");

        // The parent node where new nodes will be inserted.
        CompositeNode destinationParent = insertionDestination.ParentNode;

        // Create a NodeImporter that uses the destination's styles.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.UseDestinationStyles);

        // Loop through all block‑level nodes in the source document and import them.
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

                // Import the node (deep clone) together with its children.
                Node importedNode = importer.ImportNode(srcNode, true);

                // Insert the imported node after the current insertion point.
                destinationParent.InsertAfter(importedNode, insertionDestination);

                // Move the insertion point forward so subsequent nodes are appended sequentially.
                insertionDestination = importedNode;
            }
        }

        // Save the merged document.
        dstDoc.Save(outputPath);
    }

    // Example usage.
    public static void Main()
    {
        string destinationFile = @"C:\Docs\Destination.docx";
        string sourceFile = @"C:\Docs\Source.docx";
        string resultFile = @"C:\Docs\MergedResult.docx";

        InsertDocumentUsingNodeImporter(destinationFile, sourceFile, resultFile);
        Console.WriteLine("Document merged successfully.");
    }
}
