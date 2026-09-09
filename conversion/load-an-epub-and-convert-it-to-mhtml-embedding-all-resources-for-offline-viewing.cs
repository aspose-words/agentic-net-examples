using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document that will be saved as EPUB.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document that will be saved as EPUB and then converted to MHTML.");
        const string epubPath = "sample.epub";
        sourceDoc.Save(epubPath, SaveFormat.Epub);

        // Load the previously saved EPUB file.
        Document epubDoc = new Document(epubPath);

        // Configure save options for MHTML with all resources embedded.
        HtmlSaveOptions mhtmlOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            ExportFontResources = true,                 // Embed fonts.
            ExportCidUrlsForMhtmlResources = true       // Use CID URLs for resources.
        };

        const string mhtmlPath = "output.mht";
        epubDoc.Save(mhtmlPath, mhtmlOptions);

        // Verify that the MHTML file was created.
        if (!File.Exists(mhtmlPath))
            throw new InvalidOperationException("The MHTML output file was not created.");

        // Optional: clean up temporary files.
        // File.Delete(epubPath);
    }
}
