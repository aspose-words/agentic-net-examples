using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // ----- Create primary header -----
        HeaderFooter header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
        doc.FirstSection.HeadersFooters.Add(header);
        // Add a paragraph to the header.
        header.AppendParagraph("Primary header");

        // ----- Create primary footer -----
        HeaderFooter footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
        doc.FirstSection.HeadersFooters.Add(footer);
        // Add a paragraph to the footer.
        footer.AppendParagraph("Primary footer");

        // ----- Build the main body with three pages -----
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Write("Page 3");

        // ----- Configure TXT save options to include headers and footers -----
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            // Export only primary headers/footers at the beginning and end of each section.
            ExportHeadersFootersMode = TxtExportHeadersFootersMode.PrimaryOnly
        };

        // ----- Save the document as plain text -----
        string txtPath = "Output.txt";
        doc.Save(txtPath, saveOptions);

        // Optional: output the resulting text to the console.
        Console.WriteLine(File.ReadAllText(txtPath));
    }
}
