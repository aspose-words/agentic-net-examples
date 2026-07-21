using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for temporary files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // -----------------------------------------------------------------
        // 1. Create the destination document that contains a bookmark.
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        destBuilder.Writeln("Start of destination document.");
        destBuilder.StartBookmark("InsertHere");
        destBuilder.Writeln("<<Bookmark location>>");
        destBuilder.EndBookmark("InsertHere");
        destBuilder.Writeln("End of destination document.");

        // -----------------------------------------------------------------
        // 2. Create the source document that contains several paragraphs
        //    and a table (the table will be ignored during import).
        // -----------------------------------------------------------------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);

        srcBuilder.Writeln("Source paragraph 1.");
        srcBuilder.Writeln("Source paragraph 2.");

        // Insert a table – this node must not be imported.
        srcBuilder.StartTable();
        srcBuilder.InsertCell();
        srcBuilder.Write("Cell 1");
        srcBuilder.EndRow();
        srcBuilder.EndTable();

        srcBuilder.Writeln("Source paragraph after table.");

        // Save the source document to disk (optional, just to demonstrate file I/O).
        string srcPath = Path.Combine(workDir, "Source.docx");
        srcDoc.Save(srcPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 3. Locate the bookmark in the destination document.
        // -----------------------------------------------------------------
        Bookmark bookmark = destDoc.Range.Bookmarks["InsertHere"];
        if (bookmark == null)
            throw new InvalidOperationException("Bookmark 'InsertHere' was not found.");

        // The insertion point is the paragraph that contains the bookmark start.
        Node insertionNode = bookmark.BookmarkStart.ParentNode;

        // -----------------------------------------------------------------
        // 4. Import only paragraph nodes from the source document after the bookmark.
        // -----------------------------------------------------------------
        InsertParagraphsOnly(insertionNode, srcDoc);

        // -----------------------------------------------------------------
        // 5. Save the merged document.
        // -----------------------------------------------------------------
        string resultPath = Path.Combine(workDir, "Merged.docx");
        destDoc.Save(resultPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 6. Simple validation – ensure the file exists and contains the expected text.
        // -----------------------------------------------------------------
        if (!File.Exists(resultPath))
            throw new FileNotFoundException("Merged document was not created.", resultPath);

        Document resultDoc = new Document(resultPath);
        string resultText = resultDoc.GetText();

        // Expected paragraphs from the source document (the table should be absent).
        if (!resultText.Contains("Source paragraph 1.") ||
            !resultText.Contains("Source paragraph 2.") ||
            !resultText.Contains("Source paragraph after table.") ||
            resultText.Contains("Cell 1"))
        {
            throw new InvalidOperationException("The merged document does not contain the expected content.");
        }

        // Output a short confirmation (no interactive input required).
        Console.WriteLine("Document merged successfully. Output saved to: " + resultPath);
    }

    // Inserts only paragraph nodes from srcDoc after insertionDestination.
    private static void InsertParagraphsOnly(Node insertionDestination, Document srcDoc)
    {
        // The destination must be a paragraph or a table.
        if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
            throw new ArgumentException("The insertion destination must be a paragraph or a table.");

        CompositeNode destinationParent = insertionDestination.ParentNode;

        // NodeImporter handles style and list translation.
        NodeImporter importer = new NodeImporter(srcDoc, insertionDestination.Document, ImportFormatMode.KeepSourceFormatting);

        // Iterate over all block-level nodes in each source section.
        foreach (Section srcSection in srcDoc.Sections)
        {
            foreach (Node srcNode in srcSection.Body)
            {
                // Process only paragraphs.
                if (srcNode.NodeType != NodeType.Paragraph)
                    continue;

                Paragraph para = (Paragraph)srcNode;

                // Skip the last empty paragraph of a section (Aspose.Words adds it automatically).
                if (para.IsEndOfSection && !para.HasChildNodes)
                    continue;

                // Import the paragraph into the destination document.
                Node importedNode = importer.ImportNode(para, true);
                destinationParent.InsertAfter(importedNode, insertionDestination);
                insertionDestination = importedNode; // Move the insertion point forward.
            }
        }
    }
}
