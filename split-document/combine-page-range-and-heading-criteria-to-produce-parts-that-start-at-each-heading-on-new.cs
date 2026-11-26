using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Layout;

public class Program
{
    public static void Main()
    {
        // ------------------------------------------------------------
        // 1. Create a sample document with several Heading1 sections.
        // ------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Helper that adds a Heading1 followed by a number of normal paragraphs.
        void AddHeading(string text, int paragraphsAfter)
        {
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln(text);
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            for (int i = 0; i < paragraphsAfter; i++)
            {
                builder.Writeln($"Paragraph {i + 1} after {text}. Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                                 "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
            }
        }

        // Build the document – each heading will span several pages.
        AddHeading("Chapter 1", 30);
        AddHeading("Chapter 2", 25);
        AddHeading("Chapter 3", 35);
        AddHeading("Chapter 4", 20);

        // Save the source document (optional, for manual inspection).
        string sourcePath = "SourceDocument.docx";
        sourceDoc.Save(sourcePath);

        // ------------------------------------------------------------
        // 2. Map each Heading1 paragraph to its start page number.
        // ------------------------------------------------------------
        LayoutCollector collector = new LayoutCollector(sourceDoc);
        List<Paragraph> headingParagraphs = new List<Paragraph>();

        foreach (Paragraph para in sourceDoc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1)
                headingParagraphs.Add(para);
        }

        // Store the 1‑based start page for each heading.
        List<int> headingStartPages = new List<int>();
        foreach (Paragraph heading in headingParagraphs)
        {
            int page = collector.GetStartPageIndex(heading); // 1‑based
            headingStartPages.Add(page);
        }

        // ------------------------------------------------------------
        // 3. Split the document – each part starts at a heading that
        //    begins on a *new* page (different from the previous heading).
        // ------------------------------------------------------------
        int totalPages = sourceDoc.PageCount;
        int partIndex = 0;

        for (int i = 0; i < headingParagraphs.Count; i++)
        {
            int startPage = headingStartPages[i];

            // Skip headings that do NOT start on a new page.
            if (i > 0 && startPage == headingStartPages[i - 1])
                continue;

            // Determine the end page for this part.
            int endPage;
            if (i + 1 < headingStartPages.Count)
                endPage = headingStartPages[i + 1] - 1;
            else
                endPage = totalPages;

            // Guard against invalid ranges.
            if (startPage > endPage)
                continue;

            // Convert to zero‑based indices required by ExtractPages.
            int startIndexZero = startPage - 1;
            int pageCount = endPage - startPage + 1;

            // Extract the required pages.
            Document part = sourceDoc.ExtractPages(startIndexZero, pageCount);

            // Save the split part.
            partIndex++;
            string partPath = $"Part_{partIndex}.docx";
            part.Save(partPath);

            // Simple validation that the file was created.
            if (!File.Exists(partPath))
                throw new InvalidOperationException($"Failed to create split part: {partPath}");
        }

        // ------------------------------------------------------------
        // 4. Verify that at least one part was generated.
        // ------------------------------------------------------------
        string[] createdParts = Directory.GetFiles(Directory.GetCurrentDirectory(), "Part_*.docx");
        if (createdParts.Length == 0)
            throw new InvalidOperationException("No split parts were generated.");

        // Optional cleanup can be performed here if desired.
        // foreach (var file in createdParts) File.Delete(file);
    }
}
