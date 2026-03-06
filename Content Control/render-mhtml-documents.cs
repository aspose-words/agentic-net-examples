using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class RenderMhtmlExample
{
    static void Main()
    {
        // Path to the input DOCX file.
        string inputPath = @"C:\Docs\InputDocument.docx";

        // Path to the output MHTML file.
        string outputPath = @"C:\Docs\OutputDocument.mht";

        // Load the source document.
        Document doc = new Document(inputPath);

        // Configure save options for MHTML.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // Use CID URLs for resources (images, fonts, CSS) inside the MHTML.
            ExportCidUrlsForMhtmlResources = true,

            // Optional: export fonts and images as separate resources.
            ExportFontResources = true,
            ExportImagesAsBase64 = false,

            // Optional: make the HTML output pretty-formatted.
            PrettyFormat = true
        };

        // Save the document as MHTML using the configured options.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("MHTML document saved successfully.");
            // Optionally, read the content to check for CID references.
            string mhtmlContent = File.ReadAllText(outputPath);
            Console.WriteLine("Contains CID reference: " + mhtmlContent.Contains("cid:"));
        }
        else
        {
            Console.WriteLine("Failed to save MHTML document.");
        }
    }
}
