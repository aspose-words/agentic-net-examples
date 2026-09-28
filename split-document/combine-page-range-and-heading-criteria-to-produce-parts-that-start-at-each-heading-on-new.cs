using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Layout;

public class Program
{
    public static void Main()
    {
        // Create a sample document with headings that span several pages.
        const string sourcePath = "Source.docx";
        CreateSampleDocument(sourcePath);

        // Load the source document.
        Document sourceDoc = new Document(sourcePath);

        // Collect layout information to map headings to page numbers.
        LayoutCollector layoutCollector = new LayoutCollector(sourceDoc);

        // Find all Heading 1 paragraphs.
        List<Paragraph> headingParagraphs = new List<Paragraph>();
        foreach (Paragraph para in sourceDoc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1)
                headingParagraphs.Add(para);
        }

        if (headingParagraphs.Count == 0)
            throw new InvalidOperationException("No Heading 1 paragraphs were found in the source document.");

        // Determine page ranges for each heading.
        int totalPages = sourceDoc.PageCount;
        List<(int startPageZeroBased, int pageCount)> ranges = new List<(int, int)>();

        for (int i = 0; i < headingParagraphs.Count; i++)
        {
            // LayoutCollector returns 1‑based page numbers.
            int startPage = layoutCollector.GetStartPageIndex(headingParagraphs[i]); // 1‑based
            int endPage = (i + 1 < headingParagraphs.Count)
                ? layoutCollector.GetStartPageIndex(headingParagraphs[i + 1]) - 1
                : totalPages;

            // Guard against incorrect ordering.
            if (endPage < startPage)
                endPage = startPage;

            // Convert to zero‑based start index required by ExtractPages.
            int startZero = startPage - 1;
            int count = endPage - startPage + 1;

            // Ensure we never request pages beyond the document.
            if (startZero + count > totalPages)
                count = totalPages - startZero;

            ranges.Add((startZero, count));
        }

        // Extract each range into a separate document.
        for (int i = 0; i < ranges.Count; i++)
        {
            var (startZero, count) = ranges[i];
            Document part = sourceDoc.ExtractPages(startZero, count);
            string partPath = $"Part_{i + 1}.docx";
            part.Save(partPath);
        }

        // Validate that the expected split files exist.
        for (int i = 0; i < ranges.Count; i++)
        {
            string partPath = $"Part_{i + 1}.docx";
            if (!File.Exists(partPath))
                throw new FileNotFoundException($"Expected split file not found: {partPath}");
        }

        // Optional cleanup.
        // File.Delete(sourcePath);
    }

    // Creates a sample document with three chapters, each beginning with a Heading 1.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int chapter = 1; chapter <= 3; chapter++)
        {
            // Heading 1.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln($"Chapter {chapter}");

            // Normal paragraphs to generate multiple pages.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            for (int p = 0; p < 30; p++)
            {
                builder.Writeln(
                    $"Paragraph {p + 1} of Chapter {chapter}. Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                    $"Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
            }

            // Page break before the next chapter (except after the last one).
            if (chapter < 3)
                builder.InsertBreak(BreakType.PageBreak);
        }

        doc.Save(filePath);
    }
}
