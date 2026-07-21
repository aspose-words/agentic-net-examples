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

        // Add sample footers of each type to the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Primary footer");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterFirst);
        builder.Writeln("First page footer");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterEven);
        builder.Writeln("Even page footer");

        // Remove all footers from the first section.
        Section section = doc.FirstSection;

        HeaderFooter footer = section.HeadersFooters[HeaderFooterType.FooterFirst];
        footer?.Remove();

        footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
        footer?.Remove();

        footer = section.HeadersFooters[HeaderFooterType.FooterEven];
        footer?.Remove();

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);
    }
}
