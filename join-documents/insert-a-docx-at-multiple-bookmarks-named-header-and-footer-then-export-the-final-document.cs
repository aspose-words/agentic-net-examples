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

        // Paths for the temporary source document and the final HTML.
        string sourceDocPath = Path.Combine(outputDir, "Source.docx");
        string resultHtmlPath = Path.Combine(outputDir, "Result.html");

        // -----------------------------------------------------------------
        // 1. Create a source DOCX that will be inserted at the bookmarks.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("=== Inserted Content ===");
        srcBuilder.Writeln("This text comes from the source DOCX.");
        sourceDoc.Save(sourceDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Create the main document containing two bookmarks: Header and Footer.
        // -----------------------------------------------------------------
        Document mainDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(mainDoc);

        builder.Writeln("Main document start.");

        // Bookmark named Header.
        builder.StartBookmark("Header");
        builder.Writeln("[Header placeholder]");
        builder.EndBookmark("Header");

        builder.Writeln("Some intermediate content.");

        // Bookmark named Footer.
        builder.StartBookmark("Footer");
        builder.Writeln("[Footer placeholder]");
        builder.EndBookmark("Footer");

        builder.Writeln("Main document end.");

        // -----------------------------------------------------------------
        // 3. Load the source document and insert it at each bookmark.
        // -----------------------------------------------------------------
        Document insertDoc = new Document(sourceDocPath);

        // Insert at Header bookmark.
        builder.MoveToBookmark("Header");
        builder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

        // Insert at Footer bookmark.
        builder.MoveToBookmark("Footer");
        builder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // 4. Save the merged document as HTML.
        // -----------------------------------------------------------------
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        mainDoc.Save(resultHtmlPath, htmlOptions);

        // -----------------------------------------------------------------
        // 5. Simple validation to ensure the HTML file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(resultHtmlPath))
            throw new InvalidOperationException("HTML export failed: file not found.");

        // The program finishes here without waiting for user input.
    }
}
