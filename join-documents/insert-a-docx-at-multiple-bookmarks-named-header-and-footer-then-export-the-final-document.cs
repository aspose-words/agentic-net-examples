using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a source DOCX that will be inserted.
        string sourcePath = Path.Combine(artifactsDir, "Source.docx");
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("This is the inserted content.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // Create a destination document with two bookmarks: Header and Footer.
        string templatePath = Path.Combine(artifactsDir, "Template.docx");
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        destBuilder.StartBookmark("Header");
        destBuilder.Writeln("Placeholder for header.");
        destBuilder.EndBookmark("Header");

        destBuilder.Writeln(); // Add a blank line between bookmarks.

        destBuilder.StartBookmark("Footer");
        destBuilder.Writeln("Placeholder for footer.");
        destBuilder.EndBookmark("Footer");

        destDoc.Save(templatePath, SaveFormat.Docx);

        // Load the source document to be inserted.
        Document docToInsert = new Document(sourcePath);

        // Insert the source document at the Header bookmark.
        destBuilder.MoveToBookmark("Header");
        destBuilder.InsertDocument(docToInsert, ImportFormatMode.KeepSourceFormatting);

        // Insert the source document at the Footer bookmark.
        destBuilder.MoveToBookmark("Footer");
        destBuilder.InsertDocument(docToInsert, ImportFormatMode.KeepSourceFormatting);

        // Export the merged document to HTML.
        string htmlPath = Path.Combine(artifactsDir, "Result.html");
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        destDoc.Save(htmlPath, htmlOptions);

        // Validate that the HTML file was created.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("The HTML output file was not created.");

        // Optional: indicate completion.
        Console.WriteLine("Document merged and saved as HTML at: " + htmlPath);
    }
}
