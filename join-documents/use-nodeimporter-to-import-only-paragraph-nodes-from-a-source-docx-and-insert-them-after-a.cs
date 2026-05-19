using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Prepare folders and file names.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        string sourcePath = Path.Combine(outputDir, "Source.docx");
        string destinationPath = Path.Combine(outputDir, "Destination.docx");
        string mergedPath = Path.Combine(outputDir, "Merged.docx");

        // ---------- Create source document ----------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("Source Paragraph 1");
        srcBuilder.Writeln("Source Paragraph 2");
        srcBuilder.Writeln("Source Paragraph 3");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // ---------- Create destination document with a bookmark ----------
        Document destDoc = new Document();
        DocumentBuilder dstBuilder = new DocumentBuilder(destDoc);
        dstBuilder.Writeln("Destination start.");
        dstBuilder.StartBookmark("InsertHere");
        dstBuilder.Writeln("Bookmark placeholder.");
        dstBuilder.EndBookmark("InsertHere");
        dstBuilder.Writeln("Destination end.");
        destDoc.Save(destinationPath, SaveFormat.Docx);

        // ---------- Load documents ----------
        Document src = new Document(sourcePath);
        Document dst = new Document(destinationPath);

        // Locate the bookmark where content will be inserted.
        Bookmark bookmark = dst.Range.Bookmarks["InsertHere"];
        if (bookmark == null)
            throw new InvalidOperationException("Bookmark 'InsertHere' not found.");

        // The node after which we will insert imported paragraphs.
        Node insertionNode = bookmark.BookmarkStart.ParentNode;

        // Initialize NodeImporter with KeepSourceFormatting to preserve source styles.
        NodeImporter importer = new NodeImporter(src, dst, ImportFormatMode.KeepSourceFormatting);

        // Import only paragraph nodes from the source document.
        foreach (Section srcSection in src.Sections)
        {
            foreach (Node srcNode in srcSection.Body)
            {
                if (srcNode.NodeType != NodeType.Paragraph)
                    continue;

                Paragraph para = (Paragraph)srcNode;

                // Skip the last empty paragraph of a section (Word adds it automatically).
                if (para.IsEndOfSection && !para.HasChildNodes)
                    continue;

                // Import the paragraph node (deep clone) into the destination document.
                Node importedNode = importer.ImportNode(para, true);

                // Insert the imported paragraph after the current insertion point.
                CompositeNode parent = insertionNode.ParentNode;
                parent.InsertAfter(importedNode, insertionNode);
                insertionNode = importedNode; // Move insertion point forward.
            }
        }

        // Save the merged document.
        dst.Save(mergedPath, SaveFormat.Docx);

        // Simple validation that the merged file was created.
        if (!File.Exists(mergedPath))
            throw new FileNotFoundException("Merged document was not saved.", mergedPath);
    }
}
