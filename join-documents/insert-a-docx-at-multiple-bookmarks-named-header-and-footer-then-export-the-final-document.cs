using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary working folder.
        string workFolder = Path.Combine(Path.GetTempPath(), "AsposeJoinExample");
        Directory.CreateDirectory(workFolder);

        // File paths for the main document, the document to insert, and the final HTML output.
        string mainDocPath = Path.Combine(workFolder, "Main.docx");
        string insertDocPath = Path.Combine(workFolder, "Insert.docx");
        string outputHtmlPath = Path.Combine(workFolder, "Result.html");

        // -------------------------------------------------------------
        // Create the main document containing two bookmarks: Header and Footer.
        // -------------------------------------------------------------
        Document mainDoc = new Document();
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);

        mainBuilder.Writeln("Document with bookmarks.");

        // Header bookmark.
        mainBuilder.StartBookmark("Header");
        mainBuilder.Writeln("Header placeholder.");
        mainBuilder.EndBookmark("Header");

        // Page break between the two sections.
        mainBuilder.InsertBreak(BreakType.PageBreak);

        // Footer bookmark.
        mainBuilder.StartBookmark("Footer");
        mainBuilder.Writeln("Footer placeholder.");
        mainBuilder.EndBookmark("Footer");

        // Save the main document as DOCX.
        mainDoc.Save(mainDocPath, SaveFormat.Docx);

        // -------------------------------------------------------------
        // Create the document that will be inserted at each bookmark.
        // -------------------------------------------------------------
        Document insertDoc = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("Inserted content.");
        insertDoc.Save(insertDocPath, SaveFormat.Docx);

        // -------------------------------------------------------------
        // Load the documents for processing.
        // -------------------------------------------------------------
        Document destination = new Document(mainDocPath);
        Document source = new Document(insertDocPath);

        // Insert the source document at each bookmark.
        foreach (string bookmarkName in new[] { "Header", "Footer" })
        {
            // Ensure the bookmark exists.
            bool bookmarkExists = destination.Range.Bookmarks.Any(b => b.Name == bookmarkName);
            if (!bookmarkExists)
                throw new InvalidOperationException($"Bookmark '{bookmarkName}' not found.");

            // Clear any existing placeholder text inside the bookmark.
            Bookmark bm = destination.Range.Bookmarks[bookmarkName];
            bm.Text = string.Empty;

            // Move the builder to the bookmark location and insert the source document.
            DocumentBuilder builder = new DocumentBuilder(destination);
            builder.MoveToBookmark(bookmarkName);
            builder.InsertDocument(source, ImportFormatMode.KeepSourceFormatting);
        }

        // -------------------------------------------------------------
        // Export the final merged document to HTML.
        // -------------------------------------------------------------
        destination.Save(outputHtmlPath, SaveFormat.Html);

        // Validate that the HTML file was created.
        if (!File.Exists(outputHtmlPath))
            throw new FileNotFoundException("The HTML output file was not created.", outputHtmlPath);
    }
}
