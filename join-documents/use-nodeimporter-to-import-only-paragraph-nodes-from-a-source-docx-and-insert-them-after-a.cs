using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // File names in the application folder.
        const string destPath = "Destination.docx";
        const string srcPath = "Source.docx";
        const string outputPath = "Merged.docx";

        // ---------- Create destination document with a bookmark ----------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("This is the beginning of the destination document.");
        destBuilder.StartBookmark("InsertHere");
        destBuilder.Writeln("Bookmark location.");
        destBuilder.EndBookmark("InsertHere");
        destBuilder.Writeln("This is the end of the destination document.");
        destDoc.Save(destPath);

        // ---------- Create source document containing paragraphs and a table ----------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("First paragraph from source.");
        srcBuilder.Writeln("Second paragraph from source.");

        // Add a table to demonstrate that non‑paragraph nodes are ignored.
        srcBuilder.StartTable();
        srcBuilder.InsertCell();
        srcBuilder.Write("Cell 1");
        srcBuilder.InsertCell();
        srcBuilder.Write("Cell 2");
        srcBuilder.EndTable();

        srcBuilder.Writeln("Third paragraph after table.");
        srcDoc.Save(srcPath);

        // ---------- Load documents ----------
        Document destination = new Document(destPath);
        Document source = new Document(srcPath);

        // Locate the bookmark where the paragraphs will be inserted.
        Bookmark bookmark = destination.Range.Bookmarks["InsertHere"];
        if (bookmark == null)
            throw new InvalidOperationException("Bookmark 'InsertHere' not found.");

        // Insert only paragraph nodes after the bookmark's paragraph.
        InsertParagraphsAfterNode(bookmark.BookmarkStart.ParentNode, source);

        // Save the merged result.
        destination.Save(outputPath);

        // Simple validation.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("Merged document was not saved.", outputPath);
    }

    // Inserts only paragraph nodes from srcDoc after insertionDestination.
    private static void InsertParagraphsAfterNode(Node insertionDestination, Document srcDoc)
    {
        // Destination must be a paragraph or a table.
        if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
            throw new ArgumentException("The destination node must be a paragraph or a table.");

        CompositeNode destinationParent = insertionDestination.ParentNode;

        // NodeImporter translates styles, lists, etc.
        NodeImporter importer = new NodeImporter(srcDoc, insertionDestination.Document, ImportFormatMode.KeepSourceFormatting);

        foreach (Section srcSection in srcDoc.Sections)
        {
            foreach (Node srcNode in srcSection.Body)
            {
                // Process only paragraphs.
                if (srcNode.NodeType != NodeType.Paragraph)
                    continue;

                Paragraph para = (Paragraph)srcNode;

                // Skip the final empty paragraph that Aspose.Words adds to each section.
                if (para.IsEndOfSection && !para.HasChildNodes)
                    continue;

                Node importedNode = importer.ImportNode(srcNode, true);
                destinationParent.InsertAfter(importedNode, insertionDestination);
                insertionDestination = importedNode;
            }
        }
    }
}
