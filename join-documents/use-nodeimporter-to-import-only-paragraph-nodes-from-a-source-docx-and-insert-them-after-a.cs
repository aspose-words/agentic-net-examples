using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the destination document with a bookmark named "InsertHere".
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        destBuilder.Writeln("Paragraph before the bookmark.");
        destBuilder.StartBookmark("InsertHere");
        destBuilder.Writeln("Bookmark placeholder.");
        destBuilder.EndBookmark("InsertHere");
        destBuilder.Writeln("Paragraph after the bookmark.");

        // Create the source document that contains only paragraphs to be imported.
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("First imported paragraph.");
        srcBuilder.Writeln("Second imported paragraph.");
        srcBuilder.Writeln("Third imported paragraph.");

        // Locate the bookmark in the destination document.
        Bookmark bookmark = destDoc.Range.Bookmarks["InsertHere"];
        if (bookmark == null)
            throw new InvalidOperationException("Bookmark 'InsertHere' not found.");

        // The node after which the imported paragraphs will be inserted.
        Node insertionNode = bookmark.BookmarkStart.ParentNode;

        // Import only paragraph nodes from the source document.
        InsertParagraphsOnly(insertionNode, srcDoc);

        // Save the resulting document.
        string outputPath = "Result.docx";
        destDoc.Save(outputPath, SaveFormat.Docx);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The merged document was not saved.", outputPath);
    }

    // Inserts only paragraph nodes from srcDoc after insertionDestination.
    private static void InsertParagraphsOnly(Node insertionDestination, Document srcDoc)
    {
        // The destination node must be a paragraph or a table.
        if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
            throw new ArgumentException("The destination node must be a paragraph or a table.");

        CompositeNode destinationParent = insertionDestination.ParentNode;

        // NodeImporter efficiently imports nodes while preserving styles and lists.
        NodeImporter importer = new NodeImporter(srcDoc, insertionDestination.Document, ImportFormatMode.KeepSourceFormatting);

        // Iterate over all sections and their body nodes in the source document.
        foreach (Section srcSection in srcDoc.Sections)
        {
            foreach (Node srcNode in srcSection.Body)
            {
                // Process only paragraph nodes.
                if (srcNode.NodeType != NodeType.Paragraph)
                    continue;

                Paragraph para = (Paragraph)srcNode;

                // Skip the last empty paragraph of a section (Word adds it automatically).
                if (para.IsEndOfSection && !para.HasChildNodes)
                    continue;

                // Import the paragraph into the destination document.
                Node importedNode = importer.ImportNode(para, true);

                // Insert the imported paragraph after the current insertion point.
                destinationParent.InsertAfter(importedNode, insertionDestination);
                insertionDestination = importedNode;
            }
        }
    }
}
