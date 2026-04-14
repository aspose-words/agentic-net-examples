using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string workDir = Directory.GetCurrentDirectory();
        string outputDir = Path.Combine(workDir, "Output");
        Directory.CreateDirectory(outputDir);

        // File names
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        string resultPath = Path.Combine(outputDir, "Result.odt");

        // -----------------------------------------------------------------
        // 1. Create the source document that will be inserted at the bookmark.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is the content of the inserted DOCX document.");
        srcBuilder.Writeln("It will appear inside a table cell of the main document.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Create the destination document with a table and a bookmark inside a cell.
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        // Start a table.
        destBuilder.StartTable();

        // First cell – contains the bookmark where the source document will be inserted.
        destBuilder.InsertCell();
        destBuilder.StartBookmark("InsertHere");
        destBuilder.Write("Placeholder before insertion. ");
        destBuilder.EndBookmark("InsertHere");

        // Second cell – just for demonstration.
        destBuilder.InsertCell();
        destBuilder.Write("Second cell content.");

        // End the row and the table.
        destBuilder.EndRow();
        destBuilder.EndTable();

        // -----------------------------------------------------------------
        // 3. Load the source document (already in memory) and insert it at the bookmark.
        // -----------------------------------------------------------------
        // Move the cursor to the bookmark.
        destBuilder.MoveToBookmark("InsertHere");

        // Insert the source document inline, preserving its formatting.
        destBuilder.InsertDocumentInline(sourceDoc, ImportFormatMode.KeepSourceFormatting, new ImportFormatOptions());

        // -----------------------------------------------------------------
        // 4. Save the merged document as ODT, preserving the table structure.
        // -----------------------------------------------------------------
        OdtSaveOptions odtOptions = new OdtSaveOptions();
        destDoc.Save(resultPath, odtOptions);

        // -----------------------------------------------------------------
        // 5. Validation – ensure the file exists and contains expected text.
        // -----------------------------------------------------------------
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The ODT file was not created.");

        // Load the saved ODT to verify its content.
        Document verificationDoc = new Document(resultPath);
        string text = verificationDoc.GetText();

        if (!text.Contains("This is the content of the inserted DOCX document.") ||
            !text.Contains("Second cell content."))
        {
            throw new InvalidOperationException("The merged content was not found in the output document.");
        }

        // Optional: indicate success (no interactive prompts required).
        Console.WriteLine("Document merged and saved successfully to: " + resultPath);
    }
}
