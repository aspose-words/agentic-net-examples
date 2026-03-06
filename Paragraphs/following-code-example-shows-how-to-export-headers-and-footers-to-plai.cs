using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class ExportHeadersFootersToPlainText
{
    static void Main()
    {
        // Path to the folder where output files will be saved.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();

        // Add even and primary headers/footers.
        // Even header.
        HeaderFooter evenHeader = new HeaderFooter(doc, HeaderFooterType.HeaderEven);
        evenHeader.AppendParagraph("Even header");
        doc.FirstSection.HeadersFooters.Add(evenHeader);

        // Even footer.
        HeaderFooter evenFooter = new HeaderFooter(doc, HeaderFooterType.FooterEven);
        evenFooter.AppendParagraph("Even footer");
        doc.FirstSection.HeadersFooters.Add(evenFooter);

        // Primary header.
        HeaderFooter primaryHeader = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
        primaryHeader.AppendParagraph("Primary header");
        doc.FirstSection.HeadersFooters.Add(primaryHeader);

        // Primary footer.
        HeaderFooter primaryFooter = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
        primaryFooter.AppendParagraph("Primary footer");
        doc.FirstSection.HeadersFooters.Add(primaryFooter);

        // Insert three pages of body text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Write("Page 3");

        // Choose the export mode for headers/footers.
        // Change this value to TxtExportHeadersFootersMode.None,
        // TxtExportHeadersFootersMode.PrimaryOnly, or TxtExportHeadersFootersMode.AllAtEnd
        // to see the different behaviours.
        TxtExportHeadersFootersMode exportMode = TxtExportHeadersFootersMode.PrimaryOnly;

        // Configure TxtSaveOptions with the selected mode.
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            ExportHeadersFootersMode = exportMode
        };

        // Save the document as plain text using the configured options.
        string txtPath = Path.Combine(artifactsDir, "ExportHeadersFooters.txt");
        doc.Save(txtPath, saveOptions);

        // Load the saved plain‑text file and output its contents.
        string docText = File.ReadAllText(txtPath);
        Console.WriteLine("Export mode: " + exportMode);
        Console.WriteLine("Resulting text:");
        Console.WriteLine("--------------------");
        Console.WriteLine(docText);
        Console.WriteLine("--------------------");
    }
}
