using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputPath = @"C:\Docs\input.doc";

        // Path where the MHTML file will be saved.
        string outputPath = @"C:\Docs\output.mhtml";

        // Load the DOC document.
        Document doc = new Document(inputPath);

        // Set up save options for MHTML conversion.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // Use HTML5 compliance (optional, can be omitted).
            HtmlVersion = HtmlVersion.Html5,

            // Export headers and footers per section (default behavior).
            ExportHeadersFootersMode = ExportHeadersFootersMode.PerSection,

            // Embed images directly into the MHTML as Base64.
            ExportImagesAsBase64 = true,

            // Include Aspose.Words generator name in the output.
            ExportGeneratorName = true
        };

        // Save the document as MHTML using the configured options.
        doc.Save(outputPath, saveOptions);
    }
}
