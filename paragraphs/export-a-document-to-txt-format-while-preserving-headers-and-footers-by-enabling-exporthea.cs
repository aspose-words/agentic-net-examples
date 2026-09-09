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
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a primary header.
        HeaderFooter header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
        doc.FirstSection.HeadersFooters.Add(header);
        header.AppendParagraph("Primary header");

        // Add a primary footer.
        HeaderFooter footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
        doc.FirstSection.HeadersFooters.Add(footer);
        footer.AppendParagraph("Primary footer");

        // Add some body content with page breaks.
        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Write("Page 3");

        // Configure TXT save options to export headers and footers.
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            ExportHeadersFootersMode = TxtExportHeadersFootersMode.PrimaryOnly
        };

        // Define output path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ExportedWithHeadersFooters.txt");

        // Save the document as plain text with the specified options.
        doc.Save(outputPath, saveOptions);

        // Output the result path and file content.
        Console.WriteLine("Document saved to: " + outputPath);
        Console.WriteLine("Saved text content:");
        Console.WriteLine(File.ReadAllText(outputPath));
    }
}
