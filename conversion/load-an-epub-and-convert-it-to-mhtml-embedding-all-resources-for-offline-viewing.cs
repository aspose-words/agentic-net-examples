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
        // Save the document as EPUB.
        const string epubPath = "sample.epub";
        sourceDoc.Save(epubPath, SaveFormat.Epub);

        // Load the EPUB file.
        Document epubDoc = new Document(epubPath);

        // Configure save options for MHTML with embedded resources.
        HtmlSaveOptions mhtmlOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // Use Content-ID URLs to ensure all resources are embedded.
            ExportCidUrlsForMhtmlResources = true
        };

        // Save the document as MHTML.
        const string mhtmlPath = "output.mht";
        epubDoc.Save(mhtmlPath, mhtmlOptions);

        // Verify that the MHTML file was created.
        if (!File.Exists(mhtmlPath))
            throw new InvalidOperationException("The MHTML output file was not created.");

        // Optionally, indicate success.
        Console.WriteLine("EPUB successfully converted to MHTML with embedded resources.");
    }
}
