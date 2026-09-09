using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the temporary source document and the final HTML result.
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        string resultPath = Path.Combine(outputDir, "Result.html");

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX that will be inserted at the bookmarks.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is the content of the inserted document.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // ---------------------------------------------------------------
        // 2. Create the destination document containing two bookmarks.
        // ---------------------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        // Bookmark named "Header".
        destBuilder.StartBookmark("Header");
        destBuilder.Writeln("Header placeholder.");
        destBuilder.EndBookmark("Header");

        destBuilder.Writeln(); // Add a blank paragraph between bookmarks.

        // Bookmark named "Footer".
        destBuilder.StartBookmark("Footer");
        destBuilder.Writeln("Footer placeholder.");
        destBuilder.EndBookmark("Footer");

        // ---------------------------------------------------------------
        // 3. Load the source document and insert it at each bookmark.
        // ---------------------------------------------------------------
        Document insertDoc = new Document(sourcePath);

        // Insert at the "Header" bookmark.
        destBuilder.MoveToBookmark("Header");
        destBuilder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

        // Insert at the "Footer" bookmark.
        destBuilder.MoveToBookmark("Footer");
        destBuilder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

        // ---------------------------------------------------------------
        // 4. Save the merged document as HTML.
        // ---------------------------------------------------------------
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        destDoc.Save(resultPath, htmlOptions);

        // ---------------------------------------------------------------
        // 5. Simple validation to ensure the HTML file was created and
        //    contains the inserted text.
        // ---------------------------------------------------------------
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("HTML output was not created.");

        string htmlContent = File.ReadAllText(resultPath);
        if (!htmlContent.Contains("This is the content of the inserted document."))
            throw new InvalidOperationException("Inserted content not found in HTML output.");
    }
}
