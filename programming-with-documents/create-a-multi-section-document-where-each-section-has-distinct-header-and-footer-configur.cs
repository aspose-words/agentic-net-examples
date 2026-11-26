using System;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create three sections, each with its own header and footer.
        for (int i = 1; i <= 3; i++)
        {
            // Insert a section break before sections 2 and 3.
            if (i > 1)
                builder.InsertBreak(BreakType.SectionBreakNewPage);

            // ----- Header for the current section -----
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Writeln($"Header for Section {i}");

            // ----- Footer for the current section -----
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Writeln($"Footer for Section {i}");

            // Return to the main body of the document.
            builder.MoveToDocumentEnd();

            // Add some body text to the current section.
            builder.Writeln($"This is the body of section {i}.");
            builder.Writeln(); // Add an empty line for readability.
        }

        // Save the document to disk.
        string outputPath = "MultiSectionHeadersFooters.docx";
        doc.Save(outputPath);
    }
}
