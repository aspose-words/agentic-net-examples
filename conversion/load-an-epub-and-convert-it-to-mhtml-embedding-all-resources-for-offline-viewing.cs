using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Paths for the intermediate EPUB and final MHTML files.
        string epubPath = Path.Combine(artifactsDir, "sample.epub");
        string mhtmlPath = Path.Combine(artifactsDir, "sample.mht");

        // -----------------------------------------------------------------
        // 1. Create a simple document and save it as EPUB.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, this is a sample document for EPUB conversion.");

        HtmlSaveOptions epubSaveOptions = new HtmlSaveOptions(SaveFormat.Epub)
        {
            Encoding = Encoding.UTF8,
            ExportImagesAsBase64 = true,          // Embed images directly.
            ExportFontResources = true,           // Include fonts.
            ExportDocumentProperties = true       // Preserve document properties.
        };
        doc.Save(epubPath, epubSaveOptions);

        // Verify that the EPUB file was created.
        if (!File.Exists(epubPath) || new FileInfo(epubPath).Length == 0)
            throw new InvalidOperationException("Failed to create the EPUB file.");

        // -----------------------------------------------------------------
        // 2. Load the EPUB and convert it to MHTML with all resources embedded.
        // -----------------------------------------------------------------
        Document epubDoc = new Document(epubPath);

        HtmlSaveOptions mhtmlSaveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            ExportImagesAsBase64 = true,               // Embed images as Base64.
            ExportFontResources = true,                // Embed fonts.
            ExportCidUrlsForMhtmlResources = true,    // Use CID URLs for resources.
            ExportDocumentProperties = true            // Preserve document properties.
        };
        epubDoc.Save(mhtmlPath, mhtmlSaveOptions);

        // Verify that the MHTML file was created.
        if (!File.Exists(mhtmlPath) || new FileInfo(mhtmlPath).Length == 0)
            throw new InvalidOperationException("Failed to create the MHTML file.");

        // Indicate successful completion.
        Console.WriteLine("EPUB to MHTML conversion completed successfully.");
    }
}
