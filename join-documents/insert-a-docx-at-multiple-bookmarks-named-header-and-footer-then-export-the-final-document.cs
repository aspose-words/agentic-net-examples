using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Folder for all generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a source DOCX that will be inserted at the bookmarks.
        // -----------------------------------------------------------------
        string sourcePath = Path.Combine(artifactsDir, "Source.docx");
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("=== Inserted Content ===");
        srcBuilder.Writeln("This paragraph comes from the source DOCX.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Create the destination document with two bookmarks: Header and Footer.
        // -----------------------------------------------------------------
        string destinationPath = Path.Combine(artifactsDir, "Destination.docx");
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        destBuilder.Writeln("Document start.");
        destBuilder.StartBookmark("Header");
        destBuilder.Writeln("[Header placeholder]");
        destBuilder.EndBookmark("Header");

        destBuilder.Writeln("Some middle content.");

        destBuilder.StartBookmark("Footer");
        destBuilder.Writeln("[Footer placeholder]");
        destBuilder.EndBookmark("Footer");

        destBuilder.Writeln("Document end.");
        destDoc.Save(destinationPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 3. Load the source DOCX (to be inserted) and insert it at each bookmark.
        // -----------------------------------------------------------------
        Document insertDoc = new Document(sourcePath);

        // Insert at Header bookmark.
        destBuilder.MoveToBookmark("Header");
        destBuilder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

        // Insert at Footer bookmark.
        destBuilder.MoveToBookmark("Footer");
        destBuilder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // 4. Save the merged document as HTML.
        // -----------------------------------------------------------------
        string htmlPath = Path.Combine(artifactsDir, "Result.html");
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        destDoc.Save(htmlPath, htmlOptions);

        // -----------------------------------------------------------------
        // 5. Validate that the HTML file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(htmlPath))
        {
            throw new InvalidOperationException("The HTML output file was not created.");
        }

        // Optional: indicate success (no interactive input required).
        Console.WriteLine("Document merged and saved as HTML at:");
        Console.WriteLine(htmlPath);
    }
}
