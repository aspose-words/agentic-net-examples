using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class ExportDocumentToTxt
{
    public static void Main()
    {
        // Create a folder for output files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

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

        // Add some body content with page breaks.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Write("Page 3");

        // Configure TXT save options to exclude headers and footers.
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            ExportHeadersFootersMode = TxtExportHeadersFootersMode.None
        };

        // Save the document as plain text.
        string txtPath = Path.Combine(outputDir, "Document.txt");
        doc.Save(txtPath, saveOptions);

        // Optional: display the saved text content.
        string savedText = File.ReadAllText(txtPath);
        Console.WriteLine("Saved TXT content without headers/footers:");
        Console.WriteLine(savedText);
    }
}
