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

        // Add even header and footer.
        HeaderFooter headerEven = new HeaderFooter(doc, HeaderFooterType.HeaderEven);
        doc.FirstSection.HeadersFooters.Add(headerEven);
        headerEven.AppendParagraph("Even header");

        HeaderFooter footerEven = new HeaderFooter(doc, HeaderFooterType.FooterEven);
        doc.FirstSection.HeadersFooters.Add(footerEven);
        footerEven.AppendParagraph("Even footer");

        // Add primary header and footer.
        HeaderFooter headerPrimary = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
        doc.FirstSection.HeadersFooters.Add(headerPrimary);
        headerPrimary.AppendParagraph("Primary header");

        HeaderFooter footerPrimary = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
        doc.FirstSection.HeadersFooters.Add(footerPrimary);
        footerPrimary.AppendParagraph("Primary footer");

        // Add body content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Write("Page 3");

        // Configure save options to exclude headers and footers.
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            ExportHeadersFootersMode = TxtExportHeadersFootersMode.None
        };

        // Ensure output directory exists and save the document as plain text.
        string outDir = "Output";
        Directory.CreateDirectory(outDir);
        string outPath = Path.Combine(outDir, "DocumentWithoutHeadersFooters.txt");
        doc.Save(outPath, saveOptions);

        // Output the saved text to the console.
        string result = File.ReadAllText(outPath);
        Console.WriteLine("Saved text content:");
        Console.WriteLine(result);
    }
}
