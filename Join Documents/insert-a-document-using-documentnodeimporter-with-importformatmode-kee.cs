using System;
using System.IO;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a destination document and add some initial content.
        Document dstDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(dstDoc);
        builder.Writeln("Destination start.");

        // Load the source document (DOCX) that will be inserted.
        string sourcePath = "Source.docx";
        Document srcDoc = new Document(sourcePath);

        // Insert the source document after the first paragraph of the destination.
        Node insertionPoint = dstDoc.FirstSection.Body.FirstParagraph;
        InsertDocument(insertionPoint, srcDoc);

        // Save the combined document.
        dstDoc.Save("Result.docx");
    }

    // Inserts the contents of a document after the specified node using NodeImporter.
    static void InsertDocument(Node insertionDestination, Document docToInsert)
    {
        if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
            throw new ArgumentException("The destination node must be a paragraph or a table.");

        CompositeNode destinationParent = insertionDestination.ParentNode;

        // NodeImporter handles style and list translation while preserving source formatting.
        NodeImporter importer = new NodeImporter(docToInsert, insertionDestination.Document, ImportFormatMode.KeepSourceFormatting);

        // Iterate over all block‑level nodes in each section of the source document.
        foreach (Section srcSection in docToInsert.Sections)
        {
            foreach (Node srcNode in srcSection.Body)
            {
                // Skip the final empty paragraph of a section.
                if (srcNode.NodeType == NodeType.Paragraph)
                {
                    Paragraph para = (Paragraph)srcNode;
                    if (para.IsEndOfSection && !para.HasChildNodes)
                        continue;
                }

                // Import the node (deep clone) into the destination document.
                Node newNode = importer.ImportNode(srcNode, true);

                // Insert the imported node after the current insertion point.
                destinationParent.InsertAfter(newNode, insertionDestination);
                insertionDestination = newNode;
            }
        }
    }
}
