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
        builder.Writeln("This is a sample EPUB document created for conversion to MHTML.");
        // Save the document as EPUB (the input for the conversion).
        const string epubPath = "sample.epub";
        sourceDoc.Save(epubPath, SaveFormat.Epub);

        // Load the previously saved EPUB file.
        Document epubDocument = new Document(epubPath);

        // Configure save options for MHTML with all resources embedded.
        HtmlSaveOptions mhtmlOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            ExportFontResources = true,               // Embed fonts.
            ExportImagesAsBase64 = true,               // Embed images as Base64.
            ExportCidUrlsForMhtmlResources = true,    // Use CID URLs for resources.
            ExportDocumentProperties = true            // Include document properties.
        };

        // Save the EPUB document as MHTML.
        const string mhtmlPath = "output.mht";
        epubDocument.Save(mhtmlPath, mhtmlOptions);

        // Validate that the MHTML file was created and contains data.
        if (!File.Exists(mhtmlPath))
            throw new InvalidOperationException("The MHTML output file was not created.");

        FileInfo info = new FileInfo(mhtmlPath);
        if (info.Length == 0)
            throw new InvalidOperationException("The MHTML output file is empty.");

        // Optionally, clean up the temporary EPUB file.
        if (File.Exists(epubPath))
            File.Delete(epubPath);
    }
}
