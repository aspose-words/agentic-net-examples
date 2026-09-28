using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Layout;

public class Program
{
    public static void Main()
    {
        // Prepare sample source documents.
        string sourcePath1 = "Source1.docx";
        string sourcePath2 = "Source2.docx";
        CreateSampleDocument(sourcePath1, "First document", 30); // ~1-2 pages
        CreateSampleDocument(sourcePath2, "Second document", 45); // ~2-3 pages

        // Load the first document as the destination.
        Document destination = new Document(sourcePath1);

        // Append the second document to the destination.
        Document sourceToAppend = new Document(sourcePath2);
        destination.AppendDocument(sourceToAppend, ImportFormatMode.KeepSourceFormatting);

        // Save the merged document.
        string mergedPath = "Merged.docx";
        destination.Save(mergedPath, SaveFormat.Docx);

        // Validate that the merged file exists.
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException("Merged document was not created.");

        // Update layout to ensure page numbers are calculated.
        destination.UpdatePageLayout();

        // Use LayoutCollector to retrieve page numbers of sections.
        LayoutCollector collector = new LayoutCollector(destination);

        // Validate that each section starts on a later page than the previous one.
        int previousStartPage = 0;
        foreach (Section section in destination.Sections)
        {
            int startPage = collector.GetStartPageIndex(section);
            if (startPage <= previousStartPage)
                throw new InvalidOperationException($"Section start page {startPage} is not after previous start page {previousStartPage}.");

            previousStartPage = startPage;
        }

        // Optionally, output validation result (no interactive wait).
        Console.WriteLine("All sections have correct sequential page numbers.");
    }

    // Helper method to create a simple document with repeated text to generate pages.
    private static void CreateSampleDocument(string filePath, string title, int repeatLines)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Title
        builder.Writeln(title);
        builder.Writeln(new string('=', title.Length));

        // Add repeated lines to generate content.
        for (int i = 0; i < repeatLines; i++)
        {
            builder.Writeln($"Line {i + 1}: Sample content to fill the document and create pages.");
        }

        // Ensure a section break at the end (optional).
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Save the document.
        doc.Save(filePath, SaveFormat.Docx);
    }
}
