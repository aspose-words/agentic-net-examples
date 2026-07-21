using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // File paths for the sample documents.
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        string destinationPath = Path.Combine(outputDir, "Destination.docx");
        string mergedPath = Path.Combine(outputDir, "Merged.docx");

        // -------------------------------------------------
        // Create a source document containing two tables.
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

        // First table.
        srcBuilder.StartTable();
        srcBuilder.InsertCell();
        srcBuilder.Write("Src Table 1 - Cell 1");
        srcBuilder.InsertCell();
        srcBuilder.Write("Src Table 1 - Cell 2");
        srcBuilder.EndRow();
        srcBuilder.EndTable();

        // Second table.
        srcBuilder.StartTable();
        srcBuilder.InsertCell();
        srcBuilder.Write("Src Table 2 - Cell 1");
        srcBuilder.InsertCell();
        srcBuilder.Write("Src Table 2 - Cell 2");
        srcBuilder.EndRow();
        srcBuilder.EndTable();

        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -------------------------------------------------
        // Create a destination document with a bookmark where tables will be inserted.
        // -------------------------------------------------
        Document destinationDoc = new Document();
        DocumentBuilder dstBuilder = new DocumentBuilder(destinationDoc);

        dstBuilder.Writeln("Destination document start.");

        // Bookmark marks the insertion point.
        dstBuilder.StartBookmark("InsertHere");
        dstBuilder.Writeln("Insertion point.");
        dstBuilder.EndBookmark("InsertHere");

        dstBuilder.Writeln("Destination document end.");

        destinationDoc.Save(destinationPath, SaveFormat.Docx);

        // -------------------------------------------------
        // Load the documents for processing.
        // -------------------------------------------------
        Document src = new Document(sourcePath);
        Document dst = new Document(destinationPath);

        // -------------------------------------------------
        // Import the first table from the source document.
        // -------------------------------------------------
        // The table to import (first table in the source).
        Table tableToImport = src.FirstSection.Body.Tables[0];

        // NodeImporter handles style and list translation.
        NodeImporter importer = new NodeImporter(src, dst, ImportFormatMode.KeepSourceFormatting);

        // Import the table node (deep clone) into the destination document.
        Node importedTable = importer.ImportNode(tableToImport, true);

        // Locate the bookmark's start paragraph – this will be the insertion destination.
        Bookmark bookmark = dst.Range.Bookmarks["InsertHere"];
        Node insertionPoint = bookmark.BookmarkStart.ParentNode; // Paragraph node.

        // Ensure the insertion point is a valid container parent (Body).
        CompositeNode parent = insertionPoint.ParentNode;

        // Insert the imported table after the bookmark paragraph.
        parent.InsertAfter(importedTable, insertionPoint);

        // -------------------------------------------------
        // Save the merged document.
        // -------------------------------------------------
        dst.Save(mergedPath, SaveFormat.Docx);

        // -------------------------------------------------
        // Validation: ensure the file exists and contains expected content.
        // -------------------------------------------------
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException("Merged document was not created.");

        Document mergedDoc = new Document(mergedPath);
        string mergedText = mergedDoc.GetText();

        if (!mergedText.Contains("Src Table 1 - Cell 1") || !mergedText.Contains("Src Table 1 - Cell 2"))
            throw new InvalidOperationException("The expected table content was not found in the merged document.");

        // The program finishes without requiring any user interaction.
    }
}
