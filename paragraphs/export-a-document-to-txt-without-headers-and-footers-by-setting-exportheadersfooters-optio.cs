using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare an output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document.
        Document doc = new Document();

        // Build the document content, including headers and footers.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a primary header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Primary header");

        // Add a primary footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Primary footer");

        // Return to the main body and add some pages.
        builder.MoveToDocumentEnd();
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
        string txtPath = Path.Combine(outputDir, "DocumentWithoutHeadersFooters.txt");
        doc.Save(txtPath, saveOptions);

        // Indicate completion.
        Console.WriteLine($"Document saved to: {txtPath}");
    }
}
