using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document. The format is detected automatically from the file extension.
        Document doc = new Document("input.docx");

        // Save the document as MHTML (Web archive) using the built‑in SaveFormat enumeration.
        doc.Save("output.mht", SaveFormat.Mhtml);

        // Optional: use HtmlSaveOptions for additional MHTML settings (e.g., CID URLs for resources).
        // HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Mhtml)
        // {
        //     ExportCidUrlsForMhtmlResources = true
        // };
        // doc.Save("output_with_cid.mht", options);
    }
}
