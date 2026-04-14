using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for the output files.
        const string outputPath = "Result.docx";

        // ---------- Create the destination document with a header containing a bookmark ----------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        // Move the cursor to the primary header of the first section.
        destBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Insert a paragraph that will hold the bookmark.
        destBuilder.StartBookmark("InsertDoc");
        destBuilder.Write("Header placeholder");
        destBuilder.EndBookmark("InsertDoc");

        // ---------- Create the source document that will be inserted ----------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("Inserted content line 1");
        srcBuilder.Writeln("Inserted content line 2");

        // ---------- Insert the source document at the bookmark inside the header ----------
        // Move back to the header and then to the bookmark.
        destBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        destBuilder.MoveToBookmark("InsertDoc");

        // Insert the source document preserving its formatting.
        destBuilder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // ---------- Save the merged document ----------
        destDoc.Save(outputPath, SaveFormat.Docx);

        // ---------- Validation ----------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output file was not created.");

        // Load the saved document to verify the inserted content exists.
        Document verificationDoc = new Document(outputPath);
        string text = verificationDoc.GetText();

        if (!text.Contains("Inserted content line 1") || !text.Contains("Inserted content line 2"))
            throw new InvalidOperationException("The inserted content was not found in the resulting document.");
    }
}
