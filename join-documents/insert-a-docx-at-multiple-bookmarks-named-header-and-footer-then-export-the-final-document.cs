using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create the destination document with two bookmarks: Header and Footer.
        Document destination = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destination);

        destBuilder.Writeln("Document start.");

        destBuilder.StartBookmark("Header");
        destBuilder.Writeln("Header placeholder.");
        destBuilder.EndBookmark("Header");

        destBuilder.Writeln("Middle content.");

        destBuilder.StartBookmark("Footer");
        destBuilder.Writeln("Footer placeholder.");
        destBuilder.EndBookmark("Footer");

        destBuilder.Writeln("Document end.");

        // Create the source document that will be inserted at each bookmark.
        Document source = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(source);
        srcBuilder.Writeln("Inserted content.");

        // Insert the source document at the Header bookmark.
        destBuilder.MoveToBookmark("Header");
        destBuilder.InsertDocument(source, ImportFormatMode.KeepSourceFormatting);

        // Insert the source document at the Footer bookmark.
        destBuilder.MoveToBookmark("Footer");
        destBuilder.InsertDocument(source, ImportFormatMode.KeepSourceFormatting);

        // Define output paths.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string htmlPath = Path.Combine(outputDir, "Result.html");

        // Save the merged document as HTML.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        destination.Save(htmlPath, htmlOptions);

        // Validate that the HTML file was created and contains the inserted text.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("Failed to create the HTML output file.");

        string htmlContent = File.ReadAllText(htmlPath);
        if (!htmlContent.Contains("Inserted content"))
            throw new InvalidOperationException("The inserted content was not found in the HTML output.");

        // Optionally, clean up (comment out if you want to keep the files).
        // Directory.Delete(outputDir, true);
    }
}
