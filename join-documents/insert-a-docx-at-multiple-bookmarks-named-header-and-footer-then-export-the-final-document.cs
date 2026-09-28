using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        string destinationPath = Path.Combine(Directory.GetCurrentDirectory(), "Destination.docx");
        string sourcePath = Path.Combine(Directory.GetCurrentDirectory(), "Source.docx");
        string outputHtmlPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.html");

        // Create destination document with two bookmarks: Header and Footer
        Document destinationDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destinationDoc);
        destBuilder.Writeln("Document with bookmarks.");
        destBuilder.StartBookmark("Header");
        destBuilder.Writeln("Header placeholder.");
        destBuilder.EndBookmark("Header");
        destBuilder.Writeln();
        destBuilder.StartBookmark("Footer");
        destBuilder.Writeln("Footer placeholder.");
        destBuilder.EndBookmark("Footer");
        destinationDoc.Save(destinationPath, SaveFormat.Docx);

        // Create source document to be inserted
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("Inserted content from source document.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // Load documents from the saved files
        Document destDoc = new Document(destinationPath);
        Document srcDoc = new Document(sourcePath);

        // Insert the source document at each bookmark
        string[] bookmarkNames = { "Header", "Footer" };
        foreach (string bookmarkName in bookmarkNames)
        {
            Bookmark bookmark = destDoc.Range.Bookmarks[bookmarkName];
            if (bookmark == null)
                throw new InvalidOperationException($"Bookmark '{bookmarkName}' not found.");

            DocumentBuilder builder = new DocumentBuilder(destDoc);
            builder.MoveToBookmark(bookmarkName);
            builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
        }

        // Export the final document to HTML
        destDoc.Save(outputHtmlPath, SaveFormat.Html);

        // Validate that the HTML file was created
        if (!File.Exists(outputHtmlPath))
            throw new FileNotFoundException("The HTML output file was not created.", outputHtmlPath);
    }
}
