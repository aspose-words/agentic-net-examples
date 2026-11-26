using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare file paths
        string baseDir = Directory.GetCurrentDirectory();
        string insertDocPath = Path.Combine(baseDir, "InsertDoc.docx");
        string resultPath = Path.Combine(baseDir, "Result.odt");

        // -------------------------------------------------
        // 1. Create the source DOCX that will be inserted.
        // -------------------------------------------------
        Document insertDoc = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("This is the content coming from the inserted DOCX.");
        insertDoc.Save(insertDocPath, SaveFormat.Docx);

        // -------------------------------------------------
        // 2. Create the destination document with a table
        //    containing a bookmark inside a cell.
        // -------------------------------------------------
        Document mainDoc = new Document();
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);

        // Start a table with one row and one cell.
        mainBuilder.StartTable();
        mainBuilder.InsertCell();

        // Write some text before the bookmark.
        mainBuilder.Write("Before bookmark ");

        // Insert the bookmark where the other document will be placed.
        mainBuilder.StartBookmark("InsertHere");
        // Placeholder (empty) – the content will be inserted here.
        mainBuilder.EndBookmark("InsertHere");

        // Write some text after the bookmark.
        mainBuilder.Write(" After bookmark");

        // End the cell, row, and table.
        mainBuilder.EndRow();
        mainBuilder.EndTable();

        // -------------------------------------------------
        // 3. Load the source document and insert it at the bookmark.
        // -------------------------------------------------
        Document srcDoc = new Document(insertDocPath);
        mainBuilder.MoveToBookmark("InsertHere");
        mainBuilder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // -------------------------------------------------
        // 4. Save the merged document as ODT, preserving the table.
        // -------------------------------------------------
        mainDoc.Save(resultPath, SaveFormat.Odt);

        // -------------------------------------------------
        // 5. Validate that the output file was created.
        // -------------------------------------------------
        if (!File.Exists(resultPath))
        {
            throw new InvalidOperationException("The ODT output file was not created.");
        }

        // Optional: Validate that the table still exists.
        int tableCount = mainDoc.GetChildNodes(NodeType.Table, true).Count;
        if (tableCount == 0)
        {
            throw new InvalidOperationException("The resulting document does not contain the expected table.");
        }
    }
}
