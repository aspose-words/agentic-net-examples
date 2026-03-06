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
            // Path to the source document (any format supported by Aspose.Words)
            string inputPath = @"C:\Docs\SourceDocument.docx";

            // Path where the MHTML file will be saved
            string outputPath = @"C:\Docs\ConvertedDocument.mht";

            // Load the document using the standard Document constructor (lifecycle rule)
            Document doc = new Document(inputPath);

            // Create HtmlSaveOptions for MHTML output (factory rule)
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
            {
                // Use CID URLs for resources – improves compatibility with some mail clients
                ExportCidUrlsForMhtmlResources = true,

                // Export fonts along with the document
                ExportFontResources = true,

                // Export CSS as an external stylesheet (instead of inline)
                CssStyleSheetType = CssStyleSheetType.External,

                // Make the output more readable
                PrettyFormat = true,

                // Optionally embed the generator name (Aspose.Words version) in the output
                ExportGeneratorName = true
            };

            // Save the document as MHTML using the Save(string, SaveOptions) overload (feature rule)
            doc.Save(outputPath, saveOptions);
        }
    }
}
