using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a simple document with two pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Content of page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Content of page 2.");

        const string sourceFile = "Sample.docx";
        doc.Save(sourceFile);

        // Ensure layout information (page count) is up‑to‑date.
        doc.UpdatePageLayout();

        // Split the document by individual pages using ExtractPages.
        // ExtractPages uses a zero‑based start page index.
        int totalPages = doc.PageCount;
        for (int i = 0; i < totalPages; i++)
        {
            Document pageDoc = doc.ExtractPages(i, 1);
            string pageFile = $"Sample_page{i + 1}.docx";
            pageDoc.Save(pageFile);
            Console.WriteLine($"Saved page {i + 1} to '{pageFile}'.");
        }

        // Simple validation that the split files exist.
        for (int i = 0; i < totalPages; i++)
        {
            string pageFile = $"Sample_page{i + 1}.docx";
            if (!File.Exists(pageFile))
                throw new InvalidOperationException($"Expected split file '{pageFile}' was not created.");
        }

        Console.WriteLine("Split mode set to: individual pages (extracted via ExtractPages).");
    }
}
