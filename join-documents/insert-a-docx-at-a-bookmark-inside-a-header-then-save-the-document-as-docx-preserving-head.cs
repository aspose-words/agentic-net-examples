using System;
using System.IO;
using System.Drawing;               // For Color
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -------------------------------------------------
        // 1. Create the source document (DOCX) that will be inserted.
        // -------------------------------------------------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Font.Size = 14;
        srcBuilder.Font.Color = Color.Blue;
        srcBuilder.Writeln("This is the inserted document content.");

        // Save the source document (optional, just for inspection).
        string srcPath = Path.Combine(outputDir, "Source.docx");
        srcDoc.Save(srcPath, SaveFormat.Docx);

        // -------------------------------------------------
        // 2. Create the destination document with a header containing a bookmark.
        // -------------------------------------------------
        Document dstDoc = new Document();
        DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);

        // Ensure the document has a header and place a bookmark inside it.
        dstBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        dstBuilder.StartBookmark("HeaderInsertPoint");
        dstBuilder.Writeln("Header before insertion.");
        dstBuilder.EndBookmark("HeaderInsertPoint");

        // Add some body content so the document has visible pages.
        dstBuilder.MoveToDocumentEnd();
        dstBuilder.Writeln("Main document body text.");

        // -------------------------------------------------
        // 3. Insert the source document at the bookmark inside the header.
        // -------------------------------------------------
        dstBuilder.MoveToBookmark("HeaderInsertPoint");
        dstBuilder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // -------------------------------------------------
        // 4. Save the merged document preserving header formatting.
        // -------------------------------------------------
        string resultPath = Path.Combine(outputDir, "Result.docx");
        dstDoc.Save(resultPath, SaveFormat.Docx);

        // -------------------------------------------------
        // 5. Simple validation that the file was created and contains the inserted text.
        // -------------------------------------------------
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The result document was not saved correctly.");

        HeaderFooter header = dstDoc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        string headerText = header.GetText();
        if (!headerText.Contains("This is the inserted document content."))
            throw new InvalidOperationException("The source content was not inserted into the header.");
    }
}
