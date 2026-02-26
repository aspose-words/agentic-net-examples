using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path where the output files will be written.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new empty document.
        Document doc = new Document();

        // Add even and primary headers/footers.
        // Even header.
        HeaderFooter headerEven = new HeaderFooter(doc, HeaderFooterType.HeaderEven);
        doc.FirstSection.HeadersFooters.Add(headerEven);
        headerEven.AppendParagraph("Even header");

        // Even footer.
        HeaderFooter footerEven = new HeaderFooter(doc, HeaderFooterType.FooterEven);
        doc.FirstSection.HeadersFooters.Add(footerEven);
        footerEven.AppendParagraph("Even footer");

        // Primary header.
        HeaderFooter headerPrimary = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
        doc.FirstSection.HeadersFooters.Add(headerPrimary);
        headerPrimary.AppendParagraph("Primary header");

        // Primary footer.
        HeaderFooter footerPrimary = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
        doc.FirstSection.HeadersFooters.Add(footerPrimary);
        footerPrimary.AppendParagraph("Primary footer");

        // Insert body text with page breaks to demonstrate header/footer placement.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Write("Page 3");

        // Choose the export mode for headers and footers.
        // Options: TxtExportHeadersFootersMode.None,
        //          TxtExportHeadersFootersMode.PrimaryOnly,
        //          TxtExportHeadersFootersMode.AllAtEnd
        TxtExportHeadersFootersMode exportMode = TxtExportHeadersFootersMode.PrimaryOnly;

        // Configure text save options.
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            ExportHeadersFootersMode = exportMode
        };

        // Save the document as plain text using the configured options.
        string txtPath = Path.Combine(artifactsDir, "ExportHeadersFooters.txt");
        doc.Save(txtPath, saveOptions);

        // Load the saved plain‑text file and display its contents.
        string plainText = File.ReadAllText(txtPath);
        Console.WriteLine("=== Exported Text ===");
        Console.WriteLine(plainText);
    }
}
