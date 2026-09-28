using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document with two pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Content of page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Content of page 2.");

        // Save the source document.
        string sourcePath = "Sample.docx";
        doc.Save(sourcePath);

        // Reload the document to ensure layout information is up‑to‑date.
        doc = new Document(sourcePath);
        doc.UpdatePageLayout();

        // Verify the source document was saved.
        if (!File.Exists(sourcePath))
            throw new FileNotFoundException("Source document was not saved.", sourcePath);

        // Split the document into individual pages using ExtractPages.
        int pageCount = doc.PageCount;
        for (int page = 1; page <= pageCount; page++)
        {
            // Extract a single page. ExtractPages uses zero‑based start index and a page count.
            Document pageDoc = doc.ExtractPages(page - 1, 1);

            // Save the extracted page.
            string pagePath = $"Page_{page}.docx";
            pageDoc.Save(pagePath);

            // Verify the page file was created.
            if (!File.Exists(pagePath))
                throw new FileNotFoundException($"Page {page} document was not saved.", pagePath);
        }

        // Write a simple report confirming the split operation.
        string reportPath = "SplitReport.txt";
        File.WriteAllText(reportPath,
            $"Source document: {sourcePath}{Environment.NewLine}" +
            $"Total pages split: {pageCount}{Environment.NewLine}" +
            $"Generated files: {string.Join(", ", GetPageFileNames(pageCount))}");

        // Verify the report file exists.
        if (!File.Exists(reportPath))
            throw new FileNotFoundException("Split report was not saved.", reportPath);
    }

    private static string[] GetPageFileNames(int pageCount)
    {
        string[] names = new string[pageCount];
        for (int i = 0; i < pageCount; i++)
            names[i] = $"Page_{i + 1}.docx";
        return names;
    }
}
