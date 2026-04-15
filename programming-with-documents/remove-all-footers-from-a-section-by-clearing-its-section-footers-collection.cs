using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add three kinds of footers to the first (and only) section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterFirst);
        builder.Writeln("First page footer.");

        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Primary footer (odd pages).");

        builder.MoveToHeaderFooter(HeaderFooterType.FooterEven);
        builder.Writeln("Even page footer.");

        // Remove all footers from each section by deleting the footer nodes.
        foreach (Section section in doc.Sections)
        {
            HeaderFooter footer = section.HeadersFooters[HeaderFooterType.FooterFirst];
            footer?.Remove();

            footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
            footer?.Remove();

            footer = section.HeadersFooters[HeaderFooterType.FooterEven];
            footer?.Remove();
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FootersRemoved.docx");
        doc.Save(outputPath);
    }
}
