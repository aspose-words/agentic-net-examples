using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a blank document.
        Document doc = new Document();

        // Add a primary header.
        HeaderFooter header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
        doc.FirstSection.HeadersFooters.Add(header);
        header.AppendParagraph("Primary header");

        // Add a primary footer.
        HeaderFooter footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
        doc.FirstSection.HeadersFooters.Add(footer);
        footer.AppendParagraph("Primary footer");

        // Build the body of the document with a few pages.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3");

        // Configure TXT save options to export headers and footers.
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            ExportHeadersFootersMode = TxtExportHeadersFootersMode.PrimaryOnly,
            ForcePageBreaks = true // Preserve page breaks in the plain‑text output.
        };

        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(artifactsDir);
        string txtPath = Path.Combine(artifactsDir, "DocumentWithHeadersFooters.txt");

        // Save the document as plain text with the specified options.
        doc.Save(txtPath, saveOptions);

        // Optional: display the saved text content.
        string savedText = File.ReadAllText(txtPath);
        Console.WriteLine("Saved TXT content:");
        Console.WriteLine(savedText);
    }
}
