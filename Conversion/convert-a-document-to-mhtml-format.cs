using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Input document path – can be any format supported by Aspose.Words (DOCX, DOC, RTF, etc.).
        string inputPath = @"C:\Docs\Source.docx";

        // Output path for the MHTML (Web archive) file.
        string outputPath = @"C:\Docs\Result.mht";

        // Load the source document from disk.
        Document doc = new Document(inputPath);

        // Create MHTML‑specific save options.
        // The constructor sets the SaveFormat to Mhtml.
        HtmlSaveOptions mhtmlOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // When true, resources (images, CSS, fonts) are referenced by CID URLs.
            // Set to false to use file‑name references (default behavior).
            ExportCidUrlsForMhtmlResources = false,

            // Optional: make the generated MHTML file more readable.
            PrettyFormat = true
        };

        // Save the document as MHTML using the configured options.
        doc.Save(outputPath, mhtmlOptions);

        // Alternative one‑liner (no custom options):
        // doc.Save(outputPath, SaveFormat.Mhtml);
    }
}
