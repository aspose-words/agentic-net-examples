using System;
using Aspose.Words;

namespace AsposeWordsNodeImporterDemo
{
    class Program
    {
        // Inserts the contents of srcDoc after the specified node in dstDoc using the given import mode.
        static void InsertDocument(Node insertionDestination, Document srcDoc, Document dstDoc, ImportFormatMode importMode)
        {
            // The destination node must be a paragraph or a table.
            if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
                throw new ArgumentException("The destination node should be either a paragraph or a table.");

            // Parent node that will receive the imported nodes.
            CompositeNode destinationParent = insertionDestination.ParentNode;

            // Create a NodeImporter that will handle style and list translation.
            NodeImporter importer = new NodeImporter(srcDoc, dstDoc, importMode);

            // Iterate through all block‑level nodes in each section of the source document.
            foreach (Section srcSection in srcDoc.Sections)
            {
                foreach (Node srcNode in srcSection.Body)
                {
                    // Skip the last empty paragraph of a section (Word adds it automatically).
                    if (srcNode.NodeType == NodeType.Paragraph)
                    {
                        Paragraph para = (Paragraph)srcNode;
                        if (para.IsEndOfSection && !para.HasChildNodes)
                            continue;
                    }

                    // Import the node (deep clone) and insert it after the current destination node.
                    Node newNode = importer.ImportNode(srcNode, true);
                    destinationParent.InsertAfter(newNode, insertionDestination);
                    insertionDestination = newNode; // Move the insertion point forward.
                }
            }
        }

        static void Main()
        {
            // Load the destination document (the document that will receive the imported content).
            Document dstDoc = new Document("Destination.docx");

            // Load the source document (the document whose content will be inserted).
            Document srcDoc = new Document("Source.docx");

            // Choose the node after which the source content will be inserted.
            // Here we use the first paragraph of the destination document.
            Node insertionPoint = dstDoc.FirstSection.Body.FirstParagraph;

            // -----------------------------------------------------------------
            // 1. Import using UseDestinationStyles (default Word behavior).
            // -----------------------------------------------------------------
            InsertDocument(insertionPoint, srcDoc, dstDoc, ImportFormatMode.UseDestinationStyles);
            // Save the result of the first import.
            dstDoc.Save("Result_UseDestinationStyles.docx");

            // -----------------------------------------------------------------
            // 2. Import using KeepSourceFormatting (preserve source appearance).
            // -----------------------------------------------------------------
            // Reload the original destination document to start from a clean state.
            dstDoc = new Document("Destination.docx");
            insertionPoint = dstDoc.FirstSection.Body.FirstParagraph;

            InsertDocument(insertionPoint, srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);
            dstDoc.Save("Result_KeepSourceFormatting.docx");

            // -----------------------------------------------------------------
            // 3. Import using KeepDifferentStyles (copy only differing styles).
            // -----------------------------------------------------------------
            dstDoc = new Document("Destination.docx");
            insertionPoint = dstDoc.FirstSection.Body.FirstParagraph;

            InsertDocument(insertionPoint, srcDoc, dstDoc, ImportFormatMode.KeepDifferentStyles);
            dstDoc.Save("Result_KeepDifferentStyles.docx");
        }
    }
}
