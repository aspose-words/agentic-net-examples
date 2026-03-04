using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToMhtml
{
    static void Main()
    {
        // Path to the source document.
        string inputPath = @"C:\Docs\Input.docx";

        // Path where the MHTML file will be saved.
        string outputPath = @"C:\Docs\Output.mht";

        // Load the document from disk.
        Document doc = new Document(inputPath);

        // Create save options for MHTML format and configure custom settings.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // Use CID URLs for embedded resources (images, fonts, CSS).
            ExportCidUrlsForMhtmlResources = true,

            // Export built‑in and custom document properties into the MHTML.
            ExportDocumentProperties = true,

            // Store CSS in an external file rather than inline.
            CssStyleSheetType = CssStyleSheetType.External,

            // Produce nicely indented (pretty) output.
            PrettyFormat = true,

            // Use UTF‑8 encoding without a byte order mark.
            Encoding = new UTF8Encoding(false)
        };

        // Save the document as MHTML using the configured options.
        doc.Save(outputPath, saveOptions);
    }
}
