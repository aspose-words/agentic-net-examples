using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Prepare file paths.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        string destinationPath = Path.Combine(outputDir, "Destination.docx");
        string resultPath = Path.Combine(outputDir, "Result.docx");

        // Create a source DOCX with a few paragraphs.
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("First paragraph from source.");
        srcBuilder.Writeln("Second paragraph from source.");
        srcBuilder.Writeln("Third paragraph from source.");
        srcDoc.Save(sourcePath);

        // Create a destination document that contains a bookmark.
        Document dstDoc = new Document();
        DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);
        dstBuilder.Writeln("Destination start.");
        dstBuilder.StartBookmark("InsertHere");
        dstBuilder.Writeln("Bookmark placeholder.");
        dstBuilder.EndBookmark("InsertHere");
        dstBuilder.Writeln("Destination end.");
        dstDoc.Save(destinationPath);

        // Insert only paragraph nodes from the source document after the bookmark.
        InsertParagraphsAfterBookmark(dstDoc, srcDoc, "InsertHere");

        // Save the merged result.
        dstDoc.Save(resultPath);

        // Simple validation.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("Result document was not created.");

        string resultText = dstDoc.GetText();
        if (!resultText.Contains("First paragraph from source.") ||
            !resultText.Contains("Second paragraph from source.") ||
            !resultText.Contains("Third paragraph from source."))
        {
            throw new InvalidOperationException("Source paragraphs were not inserted correctly.");
        }

        // Indicate successful completion.
        Console.WriteLine("Document merged successfully. Result saved to: " + resultPath);
    }

    /// <summary>
    /// Inserts only paragraph nodes from <paramref name="srcDoc"/> after the bookmark
    /// identified by <paramref name="bookmarkName"/> in <paramref name="dstDoc"/>.
    /// </summary>
    private static void InsertParagraphsAfterBookmark(Document dstDoc, Document srcDoc, string bookmarkName)
    {
        // Locate the bookmark.
        Bookmark bookmark = dstDoc.Range.Bookmarks[bookmarkName];
        if (bookmark == null)
            throw new ArgumentException($"Bookmark '{bookmarkName}' not found in destination document.");

        // The node after which new content will be inserted.
        Node insertionDestination = bookmark.BookmarkStart.ParentNode;

        // Parent container for the insertion (paragraph or table).
        CompositeNode destinationParent = insertionDestination.ParentNode;

        // Prepare the importer with the desired formatting mode.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);

        // Iterate through all sections of the source document.
        foreach (Section srcSection in srcDoc.Sections)
        {
            // Iterate through the body nodes of each section.
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
                Node importedNode = importer.ImportNode(srcNode, true);

                // Insert the imported paragraph after the current insertion point.
                destinationParent.InsertAfter(importedNode, insertionDestination);
                insertionDestination = importedNode;
            }
        }
    }
}
