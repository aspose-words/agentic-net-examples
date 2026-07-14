using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for temporary and final files
        string sourceDocPath = "Source.docx";
        string resultPath = "Result.odt";

        // -----------------------------------------------------------------
        // 1. Create the source DOCX that will be inserted later.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is the content of the inserted DOCX document.");
        sourceDoc.Save(sourceDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Create the destination document with a table and a bookmark.
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(destDoc);

        // Start a table.
        builder.StartTable();

        // First cell – contains a bookmark where the source document will be inserted.
        builder.InsertCell();
        builder.StartBookmark("InsertHere");
        builder.Writeln("Placeholder text before insertion.");
        builder.EndBookmark("InsertHere");

        // Second cell – just for demonstration.
        builder.InsertCell();
        builder.Writeln("Second cell content.");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // -----------------------------------------------------------------
        // 3. Load the source document and insert it at the bookmark.
        // -----------------------------------------------------------------
        Document docToInsert = new Document(sourceDocPath);
        builder.MoveToBookmark("InsertHere");
        builder.InsertDocumentInline(docToInsert, ImportFormatMode.UseDestinationStyles, new ImportFormatOptions());

        // -----------------------------------------------------------------
        // 4. Save the merged document as ODT, preserving the table structure.
        // -----------------------------------------------------------------
        destDoc.Save(resultPath, SaveFormat.Odt);

        // -----------------------------------------------------------------
        // 5. Verify that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(resultPath))
            throw new InvalidOperationException($"Failed to create the output file: {resultPath}");
    }
}
