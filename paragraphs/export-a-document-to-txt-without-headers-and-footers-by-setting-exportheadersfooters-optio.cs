using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a primary header.
        HeaderFooter header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
        doc.FirstSection.HeadersFooters.Add(header);
        header.AppendParagraph("Primary header");

        // Add a primary footer.
        HeaderFooter footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
        doc.FirstSection.HeadersFooters.Add(footer);
        footer.AppendParagraph("Primary footer");

        // Build the body of the document with three pages.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3");

        // Configure TXT save options to exclude headers and footers.
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            ExportHeadersFootersMode = TxtExportHeadersFootersMode.None
        };

        // Save the document as plain text.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.txt");
        doc.Save(outputPath, saveOptions);

        // Optional: display the saved text to verify the result.
        Console.WriteLine(File.ReadAllText(outputPath));
    }
}
