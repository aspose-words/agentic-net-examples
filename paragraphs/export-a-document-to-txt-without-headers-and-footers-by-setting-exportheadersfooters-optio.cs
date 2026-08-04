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
        header.AppendParagraph("Sample Header");

        // Add a primary footer.
        HeaderFooter footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
        doc.FirstSection.HeadersFooters.Add(footer);
        footer.AppendParagraph("Sample Footer");

        // Add body content with a page break.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First page content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page content.");

        // Configure save options to exclude headers and footers.
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            ExportHeadersFootersMode = TxtExportHeadersFootersMode.None
        };

        // Ensure the output directory exists.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Save the document as plain text without headers/footers.
        string txtPath = Path.Combine(outputDir, "DocumentWithoutHeadersFooters.txt");
        doc.Save(txtPath, saveOptions);

        // Output the resulting text to the console.
        string result = File.ReadAllText(txtPath);
        Console.WriteLine("Exported text:");
        Console.WriteLine(result);
    }
}
