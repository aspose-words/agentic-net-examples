using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        string mainDocPath = "Main.docx";
        string insertDocPath = "Insert.docx";
        string outputPath = "Result.odt";

        // -------------------------------------------------
        // 1. Create the main document with a table and a bookmark inside a cell
        // -------------------------------------------------
        Document mainDoc = new Document();
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);

        // Start a 1x1 table
        Table table = mainBuilder.StartTable();
        mainBuilder.InsertCell();

        // Add a paragraph with a bookmark named "InsertHere"
        mainBuilder.StartBookmark("InsertHere");
        mainBuilder.Writeln("Placeholder text before insertion.");
        mainBuilder.EndBookmark("InsertHere");

        // End the row and table (EndCell is optional in recent versions)
        mainBuilder.EndRow();
        mainBuilder.EndTable();

        // Save the main document (DOCX)
        mainDoc.Save(mainDocPath, SaveFormat.Docx);

        // -------------------------------------------------
        // 2. Create the document that will be inserted
        // -------------------------------------------------
        Document insertDoc = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("This is the inserted content from the DOCX document.");
        insertDoc.Save(insertDocPath, SaveFormat.Docx);

        // -------------------------------------------------
        // 3. Load the documents from disk
        // -------------------------------------------------
        Document loadedMain = new Document(mainDocPath);
        Document loadedInsert = new Document(insertDocPath);

        // -------------------------------------------------
        // 4. Insert the DOCX at the bookmark inside the table cell
        // -------------------------------------------------
        DocumentBuilder insertionBuilder = new DocumentBuilder(loadedMain);
        insertionBuilder.MoveToBookmark("InsertHere");
        insertionBuilder.InsertDocument(loadedInsert, ImportFormatMode.KeepSourceFormatting);

        // Optionally remove the bookmark after insertion
        Bookmark bm = loadedMain.Range.Bookmarks["InsertHere"];
        if (bm != null)
        {
            loadedMain.Range.Bookmarks.Remove(bm);
        }

        // -------------------------------------------------
        // 5. Save the resulting document as ODT, preserving table structure
        // -------------------------------------------------
        loadedMain.Save(outputPath, SaveFormat.Odt);

        // -------------------------------------------------
        // 6. Validate that the output file was created
        // -------------------------------------------------
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }

        // Optional cleanup of temporary files
        // File.Delete(mainDocPath);
        // File.Delete(insertDocPath);
    }
}
