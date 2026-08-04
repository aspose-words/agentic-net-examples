using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Layout;

public class Program
{
    public static void Main()
    {
        // Create sample source documents
        string doc1Path = Path.Combine(Directory.GetCurrentDirectory(), "Source1.docx");
        string doc2Path = Path.Combine(Directory.GetCurrentDirectory(), "Source2.docx");
        CreateSampleDocument(doc1Path, "Document One", 2);
        CreateSampleDocument(doc2Path, "Document Two", 3);

        // Load source documents
        Document source1 = new Document(doc1Path);
        Document source2 = new Document(doc2Path);

        // Create destination document and append sources
        Document destination = new Document();
        destination.RemoveAllChildren(); // start with an empty document

        destination.AppendDocument(source1, ImportFormatMode.KeepSourceFormatting);
        destination.AppendDocument(source2, ImportFormatMode.KeepSourceFormatting);

        // Update layout to calculate page numbers
        destination.UpdatePageLayout();

        // Validate page numbers for each section
        LayoutCollector collector = new LayoutCollector(destination);
        int previousPageNumber = 0;
        int sectionIndex = 0;
        foreach (Section section in destination.Sections)
        {
            // Get the first paragraph of the section to retrieve its page number
            Paragraph firstParagraph = section.Body.FirstParagraph;
            if (firstParagraph == null)
                throw new InvalidOperationException($"Section {sectionIndex} does not contain any paragraphs.");

            int pageNumber = collector.GetStartPageIndex(firstParagraph);
            if (pageNumber <= previousPageNumber)
                throw new InvalidOperationException($"Section {sectionIndex} starts on page {pageNumber}, which is not after previous page {previousPageNumber}.");

            previousPageNumber = pageNumber;
            sectionIndex++;
        }

        // Save merged document
        string mergedPath = Path.Combine(Directory.GetCurrentDirectory(), "Merged.docx");
        destination.Save(mergedPath, SaveFormat.Docx);

        // Validate that the merged file exists
        if (!File.Exists(mergedPath))
            throw new FileNotFoundException("Merged document was not created.", mergedPath);

        // Indicate success (no interactive input)
        Console.WriteLine("Document merge and page number validation completed successfully.");
    }

    private static void CreateSampleDocument(string path, string title, int pageCount)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln(title);
        builder.Writeln($"This document will contain {pageCount} pages.");

        for (int i = 1; i <= pageCount; i++)
        {
            builder.Writeln($"--- Page {i} content ---");
            // Add enough text to force a new page if not the last page
            if (i < pageCount)
                builder.InsertBreak(BreakType.PageBreak);
        }

        doc.Save(path, SaveFormat.Docx);
    }
}
