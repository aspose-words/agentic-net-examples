using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a blank destination document.
        Document dstDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(dstDoc);
        builder.Writeln("Destination start.");

        // Load the source document that will be inserted.
        Document srcDoc = new Document("Source.docx");

        // -------------------------------------------------
        // 1. Append the source document to the end of the destination.
        // -------------------------------------------------
        dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // -------------------------------------------------
        // 2. Insert the source document at a bookmark inside the destination.
        // -------------------------------------------------
        builder.Writeln("Before bookmark.");
        builder.StartBookmark("InsertHere");
        builder.Writeln("Bookmark placeholder.");
        builder.EndBookmark("InsertHere");
        builder.Writeln("After bookmark.");

        // Move the cursor to the bookmark and insert the document.
        builder.MoveToBookmark("InsertHere");
        builder.InsertDocument(srcDoc, ImportFormatMode.UseDestinationStyles);

        // -------------------------------------------------
        // 3. Insert the source document after a specific paragraph using NodeImporter.
        // -------------------------------------------------
        Paragraph targetParagraph = dstDoc.FirstSection.Body.FirstParagraph;
        InsertDocumentAfterNode(targetParagraph, srcDoc);

        // Save the combined document.
        dstDoc.Save("Result.docx");
    }

    // Inserts all nodes of docToInsert after the specified insertionDestination (paragraph or table).
    static void InsertDocumentAfterNode(Node insertionDestination, Document docToInsert)
    {
        if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
            throw new ArgumentException("Destination must be a paragraph or table.");

        CompositeNode parent = insertionDestination.ParentNode;
        // NodeImporter resides directly in Aspose.Words namespace; no separate Importing namespace is required.
        NodeImporter importer = new NodeImporter(docToInsert, insertionDestination.Document, ImportFormatMode.KeepSourceFormatting);

        foreach (Section srcSection in docToInsert.Sections)
        {
            foreach (Node srcNode in srcSection.Body)
            {
                // Skip the last empty paragraph of a section.
                if (srcNode.NodeType == NodeType.Paragraph)
                {
                    Paragraph para = (Paragraph)srcNode;
                    if (para.IsEndOfSection && !para.HasChildNodes)
                        continue;
                }

                Node importedNode = importer.ImportNode(srcNode, true);
                parent.InsertAfter(importedNode, insertionDestination);
                insertionDestination = importedNode;
            }
        }
    }
}
