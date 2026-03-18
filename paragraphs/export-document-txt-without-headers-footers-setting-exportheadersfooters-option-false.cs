using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ExportTxtWithoutHeadersFooters
{
    static void Main()
    {
        // Load or create a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First page content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page content.");

        // Create TxtSaveOptions and disable header/footer export.
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            ExportHeadersFootersMode = TxtExportHeadersFootersMode.None
        };

        // Save the document as plain text without headers and footers.
        string outputPath = "OutputWithoutHeadersFooters.txt";
        doc.Save(outputPath, saveOptions);

        Console.WriteLine($"Document saved to {outputPath} without headers and footers.");
    }
}
