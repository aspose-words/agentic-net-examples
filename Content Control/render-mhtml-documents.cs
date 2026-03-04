using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace MhtmlRenderingExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source document (DOCX, DOC, etc.).
            string inputPath = @"C:\Docs\SourceDocument.docx";

            // Path where the MHTML output will be saved.
            string outputPath = @"C:\Docs\RenderedDocument.mht";

            // Load the source document.
            Document doc = new Document(inputPath);

            // Configure HTML save options for MHTML output.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
            {
                // Use CID URLs for resources (images, fonts, CSS) to improve compatibility.
                ExportCidUrlsForMhtmlResources = true,

                // Export CSS as an external file.
                CssStyleSheetType = CssStyleSheetType.External,

                // Include font resources in the output.
                ExportFontResources = true,

                // Produce readable (pretty) HTML markup.
                PrettyFormat = true
            };

            // Save the document as MHTML using the configured options.
            doc.Save(outputPath, saveOptions);

            // Optional: read the generated MHTML content (e.g., for verification or further processing).
            string mhtmlContent = File.ReadAllText(outputPath);
            Console.WriteLine("MHTML document saved successfully. Size: {0} bytes", mhtmlContent.Length);
        }
    }
}
