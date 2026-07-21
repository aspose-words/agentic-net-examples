using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create the source DOCX that will be inserted later.
        // -----------------------------------------------------------------
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("Inserted line 1.");
        srcBuilder.Writeln("Inserted line 2.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Create the destination document containing a table with a bookmark.
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        // Build a simple 2‑column table.
        destBuilder.StartTable();

        // First cell – regular content.
        destBuilder.InsertCell();
        destBuilder.Write("Cell A");

        // Second cell – contains a bookmark where the source document will be placed.
        destBuilder.InsertCell();
        destBuilder.StartBookmark("InsertHere");
        destBuilder.Write("Placeholder"); // This text will be replaced.
        destBuilder.EndBookmark("InsertHere");

        destBuilder.EndRow();
        destBuilder.EndTable();

        // -----------------------------------------------------------------
        // 3. Load the source document (already in memory) and insert it at the bookmark.
        // -----------------------------------------------------------------
        destBuilder.MoveToBookmark("InsertHere");

        // Insert the source document inline, preserving its formatting.
        destBuilder.InsertDocumentInline(sourceDoc, ImportFormatMode.KeepSourceFormatting, new ImportFormatOptions());

        // -----------------------------------------------------------------
        // 4. Save the merged document as ODT, preserving the table structure.
        // -----------------------------------------------------------------
        string resultPath = Path.Combine(outputDir, "Result.odt");
        destDoc.Save(resultPath, SaveFormat.Odt);

        // -----------------------------------------------------------------
        // 5. Simple validation – ensure the file was created and the table exists.
        // -----------------------------------------------------------------
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The ODT file was not created.");

        // Reload the saved ODT to verify the table is still present.
        Document verificationDoc = new Document(resultPath);
        if (verificationDoc.FirstSection.Body.Tables.Count == 0)
            throw new InvalidOperationException("The resulting document does not contain a table.");

        // Execution completed successfully.
    }
}
