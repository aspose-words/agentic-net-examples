using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Paths to the source and destination documents.
            string srcPath = @"C:\Docs\SourceDocument.docx";
            string dstPath = @"C:\Docs\DestinationDocument.docx";
            string resultPath = @"C:\Docs\ResultDocument.docx";

            // Load the documents using the Document(string) constructor (load rule).
            Document srcDoc = new Document(srcPath);
            Document dstDoc = new Document(dstPath);

            // Choose the node after which the source content will be inserted.
            // Here we use the first paragraph of the destination document.
            Node insertionPoint = dstDoc.FirstSection.Body.FirstParagraph;

            // Insert the source document's content after the insertion point.
            InsertDocument(insertionPoint, srcDoc);

            // Save the modified destination document (save rule).
            dstDoc.Save(resultPath, SaveFormat.Docx);
        }

        /// <summary>
        /// Inserts the contents of <paramref name="docToInsert"/> after the specified <paramref name="insertionDestination"/>.
        /// Uses NodeImporter with ImportFormatMode.UseDestinationStyles.
        /// </summary>
        /// <param name="insertionDestination">Paragraph or table node after which the content will be inserted.</param>
        /// <param name="docToInsert">Document whose content will be imported.</param>
        private static void InsertDocument(Node insertionDestination, Document docToInsert)
        {
            // Ensure the destination node is a paragraph or a table.
            if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
                throw new ArgumentException("The destination node must be a Paragraph or Table.");

            // Parent node where new nodes will be inserted.
            CompositeNode destinationParent = insertionDestination.ParentNode;

            // Create a NodeImporter that will reuse destination styles.
            NodeImporter importer = new NodeImporter(
                docToInsert,                     // source document
                insertionDestination.Document,   // destination document
                ImportFormatMode.UseDestinationStyles);

            // Iterate through all block-level nodes in each source section.
            foreach (Section srcSection in docToInsert.Sections)
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

                    // Import the node (deep clone) and insert it after the current insertion point.
                    Node importedNode = importer.ImportNode(srcNode, true);
                    destinationParent.InsertAfter(importedNode, insertionDestination);
                    insertionDestination = importedNode; // Move the insertion point forward.
                }
            }
        }
    }
}
