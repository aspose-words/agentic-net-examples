using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;               // Aspose.Drawing types
using Aspose.Drawing.Imaging;       // ImageFormat

public class Program
{
    public static void Main()
    {
        // Define file names for the intermediate EPUB and final MHTML.
        const string epubPath = "sample.epub";
        const string mhtmlPath = "output.mht";

        // -----------------------------------------------------------------
        // Step 1: Create a simple Word document and save it as EPUB.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample EPUB document created for conversion.");

        // Add an image to ensure there is a resource to embed later.
        // The image is generated as a blank PNG using Aspose.Drawing.
        const string tempImagePath = "tempImage.png";
        using (var bitmap = new Bitmap(100, 100))
        {
            // Fill the bitmap with a solid color.
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightBlue);
            }

            // Save the bitmap to a temporary file.
            bitmap.Save(tempImagePath, ImageFormat.Png);
        }

        // Insert the image into the document.
        builder.InsertImage(tempImagePath);

        // Clean up the temporary image file.
        if (File.Exists(tempImagePath))
            File.Delete(tempImagePath);

        // Save the document as EPUB.
        sourceDoc.Save(epubPath, SaveFormat.Epub);

        // Verify that the EPUB file was created.
        if (!File.Exists(epubPath))
            throw new InvalidOperationException("EPUB file was not created.");

        // -----------------------------------------------------------------
        // Step 2: Load the EPUB and convert it to MHTML with all resources embedded.
        // -----------------------------------------------------------------
        Document epubDoc = new Document(epubPath);

        // Configure save options for MHTML.
        HtmlSaveOptions mhtmlOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            ExportFontResources = true,          // Embed fonts.
            ExportImagesAsBase64 = true,         // Embed images as Base64.
            ExportCidUrlsForMhtmlResources = true,
            ExportDocumentProperties = true
        };

        // Save the document as MHTML.
        epubDoc.Save(mhtmlPath, mhtmlOptions);

        // Verify that the MHTML file was created and contains data.
        if (!File.Exists(mhtmlPath) || new FileInfo(mhtmlPath).Length == 0)
            throw new InvalidOperationException("MHTML conversion failed or produced an empty file.");

        // Indicate successful completion.
        Console.WriteLine("EPUB successfully converted to MHTML with embedded resources.");
    }
}
