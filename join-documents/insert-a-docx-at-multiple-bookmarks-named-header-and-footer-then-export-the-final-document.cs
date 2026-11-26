using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create the main document with two bookmarks: Header and Footer.
        Document mainDoc = new Document();
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);

        mainBuilder.Writeln("Document start.");

        // Header bookmark.
        mainBuilder.StartBookmark("Header");
        mainBuilder.Writeln("[Header placeholder]");
        mainBuilder.EndBookmark("Header");

        mainBuilder.Writeln("Some content between bookmarks.");

        // Footer bookmark.
        mainBuilder.StartBookmark("Footer");
        mainBuilder.Writeln("[Footer placeholder]");
        mainBuilder.EndBookmark("Footer");

        mainBuilder.Writeln("Document end.");

        // Create the document that will be inserted at the bookmarks.
        Document insertDoc = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("=== Inserted Content Start ===");
        insertBuilder.Writeln("This is the first line of inserted content.");
        insertBuilder.Writeln("This is the second line of inserted content.");
        insertBuilder.Writeln("=== Inserted Content End ===");

        // Insert the insertDoc at the Header bookmark.
        DocumentBuilder insertAtHeader = new DocumentBuilder(mainDoc);
        insertAtHeader.MoveToBookmark("Header");
        insertAtHeader.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

        // Insert the insertDoc at the Footer bookmark.
        DocumentBuilder insertAtFooter = new DocumentBuilder(mainDoc);
        insertAtFooter.MoveToBookmark("Footer");
        insertAtFooter.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the final document as HTML.
        string outputPath = "Result.html";
        mainDoc.Save(outputPath, SaveFormat.Html);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }

        // Optionally, you could perform additional validation here (e.g., check file size).
        Console.WriteLine($"Document successfully saved to '{outputPath}'.");
    }
}
