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
        // 1. Create a source DOCX that will be inserted at the bookmarks.
        // -----------------------------------------------------------------
        string sourcePath = Path.Combine(artifactsDir, "Source.docx");
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("=== Inserted Content ===");
        srcBuilder.Writeln("This paragraph comes from the source document.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // ---------------------------------------------------------------
        // 2. Create the main document containing two bookmarks: Header and Footer.
        // ---------------------------------------------------------------
        string mainPath = Path.Combine(artifactsDir, "Main.docx");
        Document mainDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(mainDoc);

        builder.Writeln("Main document start.");

        // Bookmark named Header.
        builder.StartBookmark("Header");
        builder.Writeln("[Header placeholder]");
        builder.EndBookmark("Header");

        builder.Writeln("Some content between the bookmarks.");

        // Bookmark named Footer.
        builder.StartBookmark("Footer");
        builder.Writeln("[Footer placeholder]");
        builder.EndBookmark("Footer");

        builder.Writeln("Main document end.");
        mainDoc.Save(mainPath, SaveFormat.Docx);

        // ---------------------------------------------------------------
        // 3. Load the source document to be inserted.
        // ---------------------------------------------------------------
        Document insertDoc = new Document(sourcePath);

        // ---------------------------------------------------------------
        // 4. Insert the source document at each bookmark.
        // ---------------------------------------------------------------
        // Insert at Header bookmark.
        builder.MoveToBookmark("Header");
        builder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

        // Insert at Footer bookmark.
        builder.MoveToBookmark("Footer");
        builder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

        // ---------------------------------------------------------------
        // 5. Save the merged document as HTML.
        // ---------------------------------------------------------------
        string htmlPath = Path.Combine(artifactsDir, "Result.html");
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        mainDoc.Save(htmlPath, htmlOptions);

        // ---------------------------------------------------------------
        // 6. Validate that the HTML file was created.
        // ---------------------------------------------------------------
        if (!File.Exists(htmlPath))
        {
            throw new InvalidOperationException("The HTML output file was not created.");
        }

        // Optional: indicate success (no interactive input required).
        Console.WriteLine("Document merged and saved to HTML successfully.");
    }
}
