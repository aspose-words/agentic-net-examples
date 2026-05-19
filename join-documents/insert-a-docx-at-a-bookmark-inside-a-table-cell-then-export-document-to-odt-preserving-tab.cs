using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for all generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a source DOCX document that will be inserted later.
        // -----------------------------------------------------------------
        string sourceDocPath = Path.Combine(artifactsDir, "Source.docx");
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("This is the inserted document content.");
        sourceDoc.Save(sourceDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Create the destination document containing a table with a bookmark.
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        // Start a table and add a single cell.
        destBuilder.StartTable();
        destBuilder.InsertCell();

        // Define a bookmark inside the cell where the source document will be inserted.
        const string bookmarkName = "InsertHere";
        destBuilder.StartBookmark(bookmarkName);
        destBuilder.Writeln("Placeholder before insertion.");
        destBuilder.EndBookmark(bookmarkName);

        // Close the row and the table.
        destBuilder.EndRow();
        destBuilder.EndTable();

        // -----------------------------------------------------------------
        // 3. Load the source document and insert it at the bookmark.
        // -----------------------------------------------------------------
        Document srcToInsert = new Document(sourceDocPath);
        destBuilder.MoveToBookmark(bookmarkName);
        destBuilder.InsertDocument(srcToInsert, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // 4. Save the merged document as ODT, preserving the table structure.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(artifactsDir, "Merged.odt");
        destDoc.Save(outputPath, SaveFormat.Odt);

        // -----------------------------------------------------------------
        // 5. Simple validation to ensure the operation succeeded.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The merged ODT file was not created.");

        // Verify that the inserted text is present in the final document.
        string mergedText = destDoc.GetText();
        if (!mergedText.Contains("This is the inserted document content."))
            throw new InvalidOperationException("The source document content was not found in the merged output.");

        // Optional: indicate success (no interactive input required).
        Console.WriteLine("Document merged and saved successfully to: " + outputPath);
    }
}
