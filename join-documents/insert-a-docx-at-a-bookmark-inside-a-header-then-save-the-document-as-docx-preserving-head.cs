using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for the output document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");

        // ---------- Create destination document with a header containing a bookmark ----------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        // Move the builder to the primary header of the first section.
        destBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Insert a bookmark inside the header.
        destBuilder.StartBookmark("HeaderBookmark");
        destBuilder.Write("Header start. ");
        destBuilder.EndBookmark("HeaderBookmark");
        destBuilder.Writeln("Header end.");

        // ---------- Create source document that will be inserted ----------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("<<Inserted content from source DOCX>>");

        // ---------- Insert the source document at the bookmark inside the header ----------
        // Move back to the header and then to the bookmark.
        destBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        destBuilder.MoveToBookmark("HeaderBookmark");

        // Insert the source document preserving its formatting.
        destBuilder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // ---------- Save the merged document ----------
        destDoc.Save(outputPath, SaveFormat.Docx);

        // ---------- Validation ----------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output file was not created.");

        // Load the saved document and verify that the inserted text appears in the header.
        Document verifyDoc = new Document(outputPath);
        string headerText = verifyDoc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary].GetText();

        if (!headerText.Contains("Inserted content from source DOCX"))
            throw new InvalidOperationException("The inserted content was not found in the header.");

        // If execution reaches this point, the operation succeeded.
        Console.WriteLine("Document created successfully at: " + outputPath);
    }
}
