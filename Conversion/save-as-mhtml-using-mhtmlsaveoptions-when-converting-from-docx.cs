using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsMhtmlExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            string inputPath = Path.Combine("Input", "Sample.docx");

            // Path where the MHTML file will be saved.
            string outputPath = Path.Combine("Output", "Sample.mht");

            // Load the DOCX document.
            Document doc = new Document(inputPath);

            // Create HtmlSaveOptions configured for MHTML output.
            HtmlSaveOptions mhtmlOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
            {
                // Example option: use CID URLs for resources (optional).
                ExportCidUrlsForMhtmlResources = true,

                // Optional: embed CSS externally.
                CssStyleSheetType = CssStyleSheetType.External,

                // Optional: embed fonts in the MHTML.
                ExportFontResources = true,

                // Optional: make the output more readable.
                PrettyFormat = true
            };

            // Ensure the output directory exists.
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

            // Save the document as MHTML using the specified options.
            doc.Save(outputPath, mhtmlOptions);
        }
    }
}
