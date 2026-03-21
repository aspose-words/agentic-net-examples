using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a simple document to act as the source EPUB.
        Document sourceDoc = new Document();
        var builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Hello, world! This is a sample EPUB document.");

        // Save the document as a temporary EPUB file.
        string epubPath = Path.Combine(Path.GetTempPath(), "sample_input.epub");
        sourceDoc.Save(epubPath, SaveFormat.Epub);

        // Load the EPUB document we just created.
        Document doc = new Document(epubPath);

        // Configure save options for MHTML with all resources embedded.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            ExportImagesAsBase64 = true,
            ExportFontResources = true,
            ExportCidUrlsForMhtmlResources = true,
            CssStyleSheetType = CssStyleSheetType.Inline,
            PrettyFormat = true,
            Encoding = Encoding.UTF8
        };

        // Save the document as MHTML.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.mht");
        doc.Save(outputPath, saveOptions);

        Console.WriteLine($"MHTML file saved to: {outputPath}");
    }
}
