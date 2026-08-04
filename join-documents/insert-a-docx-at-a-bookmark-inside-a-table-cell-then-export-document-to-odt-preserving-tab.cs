using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for temporary source document and final ODT output.
        const string sourcePath = "Source.docx";
        const string outputPath = "Result.odt";

        // ---------- Create the source DOCX document ----------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is the content of the inserted DOCX document.");
        // Save as DOCX (optional, but ensures the file exists on disk).
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // ---------- Create the destination document with a table ----------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        // Build a simple 1x1 table.
        destBuilder.StartTable();
        destBuilder.InsertCell();

        // Insert a bookmark inside the cell where the source document will be placed.
        destBuilder.StartBookmark("InsertHere");
        destBuilder.Writeln("Placeholder text before insertion.");
        destBuilder.EndBookmark("InsertHere");

        destBuilder.EndRow();
        destBuilder.EndTable();

        // ---------- Insert the source document at the bookmark ----------
        destBuilder.MoveToBookmark("InsertHere");
        // InsertDocumentInline mimics Word's copy‑paste behavior and keeps the content inside the cell.
        destBuilder.InsertDocumentInline(sourceDoc, ImportFormatMode.KeepSourceFormatting, new ImportFormatOptions());

        // ---------- Save the merged document as ODT ----------
        destDoc.Save(outputPath, SaveFormat.Odt);

        // ---------- Simple validation ----------
        if (!File.Exists(outputPath))
        {
            throw new Exception($"The output file '{outputPath}' was not created.");
        }

        // Clean up temporary source file (optional).
        if (File.Exists(sourcePath))
        {
            File.Delete(sourcePath);
        }
    }
}
