using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with two pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Page 1 content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2 content.");

        const string sourcePath = "Sample.docx";
        doc.Save(sourcePath);

        // Ensure layout is up‑to‑date so we can obtain the page count.
        doc.UpdatePageLayout();
        int pageCount = doc.PageCount;
        Console.WriteLine($"Document contains {pageCount} page(s).");

        // Split the document into individual pages.
        for (int i = 1; i <= pageCount; i++)
        {
            // ExtractPages uses zero‑based start page index and a page count.
            Document pageDoc = doc.ExtractPages(i - 1, 1);
            string pagePath = $"Page_{i}.docx";
            pageDoc.Save(pagePath);
            Console.WriteLine($"Saved page {i} to '{pagePath}'.");
        }

        // Verify that each split file was created.
        for (int i = 1; i <= pageCount; i++)
        {
            string pagePath = $"Page_{i}.docx";
            if (!File.Exists(pagePath))
                throw new FileNotFoundException($"Expected split file not found: {pagePath}");
        }

        // Clean up temporary files.
        if (File.Exists(sourcePath))
            File.Delete(sourcePath);
        for (int i = 1; i <= pageCount; i++)
        {
            string pagePath = $"Page_{i}.docx";
            if (File.Exists(pagePath))
                File.Delete(pagePath);
        }

        Console.WriteLine("Document split into individual pages completed successfully.");
    }
}
