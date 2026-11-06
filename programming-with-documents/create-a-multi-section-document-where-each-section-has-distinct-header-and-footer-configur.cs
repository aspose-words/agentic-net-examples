using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        int totalSections = 3;

        for (int i = 1; i <= totalSections; i++)
        {
            // Insert a section break before every section except the first.
            if (i > 1)
                builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Get the index of the current (last) section.
            int sectionIndex = doc.Sections.Count - 1;

            // ----- Header -----
            builder.MoveToSection(sectionIndex);
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Writeln($"Header for Section {i}");

            // ----- Footer -----
            builder.MoveToSection(sectionIndex);
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Writeln($"Footer for Section {i}");

            // ----- Body content -----
            builder.MoveToSection(sectionIndex);
            builder.Writeln($"Content of Section {i}");
        }

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MultiSectionHeadersFooters.docx");
        doc.Save(outputPath);
    }
}
