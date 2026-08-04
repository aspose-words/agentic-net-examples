using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Paths for the documents.
        const string destPath = "Destination.docx";
        const string srcPath = "Source.docx";
        const string outputPath = "Result.docx";

        // ---------- Create destination document with a bookmark ----------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        destBuilder.Writeln("Start of destination document.");
        destBuilder.StartBookmark("InsertHere");
        destBuilder.Writeln("Bookmark location.");
        destBuilder.EndBookmark("InsertHere");
        destBuilder.Writeln("End of destination document.");

        // Save the destination document (optional, just to have a file on disk).
        destDoc.Save(destPath);

        // ---------- Create source document containing several paragraphs ----------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);

        srcBuilder.Writeln("First paragraph from source.");
        srcBuilder.Writeln("Second paragraph from source.");
        srcBuilder.Writeln("Third paragraph from source.");

        // Save the source document so that it exists on disk.
        srcDoc.Save(srcPath);

        // ---------- Import only paragraph nodes after the bookmark ----------
        Bookmark bookmark = destDoc.Range.Bookmarks["InsertHere"];
        InsertParagraphsAfterBookmark(bookmark.BookmarkStart.ParentNode, srcDoc);

        // ---------- Save the merged result ----------
        destDoc.Save(outputPath);

        // ---------- Validate that the output file was created ----------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"The file '{outputPath}' was not created.");

        // Optional: write a short confirmation to the console.
        Console.WriteLine($"Document merged successfully. Output file: {Path.GetFullPath(outputPath)}");
    }

    /// <summary>
    /// Inserts only paragraph nodes from <paramref name="srcDoc"/> after <paramref name="insertionDestination"/>.
    /// </summary>
    /// <param name="insertionDestination">A paragraph or table node after which the content will be inserted.</param>
    /// <param name="srcDoc">The source document containing paragraphs to import.</param>
    private static void InsertParagraphsAfterBookmark(Node insertionDestination, Document srcDoc)
    {
        if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
            throw new ArgumentException("The destination node must be a paragraph or a table.");

        CompositeNode destinationParent = insertionDestination.ParentNode;

        // NodeImporter handles style and list translation between documents.
        NodeImporter importer = new NodeImporter(srcDoc, insertionDestination.Document, ImportFormatMode.KeepSourceFormatting);

        // Iterate over all block-level nodes in the source document.
        foreach (Section srcSection in srcDoc.Sections)
        {
            foreach (Node srcNode in srcSection.Body)
            {
                // Process only paragraph nodes.
                if (srcNode.NodeType != NodeType.Paragraph)
                    continue;

                Paragraph para = (Paragraph)srcNode;

                // Skip the last empty paragraph of a section (it is added automatically by Aspose.Words).
                if (para.IsEndOfSection && !para.HasChildNodes)
                    continue;

                // Import the paragraph into the destination document.
                Node importedNode = importer.ImportNode(srcNode, true);

                // Insert the imported paragraph after the current insertion point.
                destinationParent.InsertAfter(importedNode, insertionDestination);
                insertionDestination = importedNode;
            }
        }
    }
}
