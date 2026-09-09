using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string destPath = Path.Combine(Directory.GetCurrentDirectory(), "Destination.docx");
        string srcPath = Path.Combine(Directory.GetCurrentDirectory(), "Source.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.odt");

        // ---------- Create the destination document with a table and a bookmark inside a cell ----------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        // Start a table.
        destBuilder.StartTable();

        // Insert the first cell.
        destBuilder.InsertCell();

        // Place a bookmark inside this cell where the source document will be inserted.
        destBuilder.StartBookmark("InsertHere");
        destBuilder.Write("Placeholder before insertion. ");
        destBuilder.EndBookmark("InsertHere");

        // End the row and the table.
        destBuilder.EndRow();
        destBuilder.EndTable();

        // Save the destination document (optional, just to have a physical file).
        destDoc.Save(destPath, SaveFormat.Docx);

        // ---------- Create the source document that will be inserted ----------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("This is the content coming from the source DOCX.");
        srcDoc.Save(srcPath, SaveFormat.Docx);

        // ---------- Load the source document (if not already loaded) ----------
        Document srcToInsert = new Document(srcPath);

        // ---------- Insert the source document at the bookmark inside the table cell ----------
        destBuilder.MoveToBookmark("InsertHere");
        destBuilder.InsertDocumentInline(srcToInsert, ImportFormatMode.KeepSourceFormatting, new ImportFormatOptions());

        // ---------- Save the merged document as ODT, preserving the table structure ----------
        destDoc.Save(outputPath, SaveFormat.Odt);

        // ---------- Validation ----------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output ODT file was not created.");

        // Load the saved ODT to verify that the source content is present.
        Document resultDoc = new Document(outputPath);
        string resultText = resultDoc.GetText();

        if (!resultText.Contains("This is the content coming from the source DOCX."))
            throw new InvalidOperationException("The source content was not found in the merged document.");

        // Indicate successful completion.
        Console.WriteLine("Document merged and saved as ODT successfully.");
    }
}
