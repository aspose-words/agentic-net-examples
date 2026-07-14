using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;          // Aspose.Drawing.Common namespace
using Aspose.Drawing.Imaging; // For ImageFormat

public class Program
{
    public static void Main()
    {
        // Define file names for the sample EPUB input and the MHTML output.
        const string epubPath = "sample.epub";
        const string mhtmlPath = "sample.mht";

        // -----------------------------------------------------------------
        // Step 1: Create a simple Word document and save it as EPUB.
        // This serves as the source EPUB file for the conversion.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document created for EPUB to MHTML conversion.");
        builder.InsertImage(ImageFromPlaceholder()); // Insert a placeholder image.
        sourceDoc.Save(epubPath, SaveFormat.Epub);

        // Verify that the EPUB file was created.
        if (!File.Exists(epubPath))
            throw new InvalidOperationException($"Failed to create the EPUB file at '{epubPath}'.");

        // -----------------------------------------------------------------
        // Step 2: Load the EPUB file.
        // -----------------------------------------------------------------
        Document epubDoc = new Document(epubPath);

        // -----------------------------------------------------------------
        // Step 3: Configure HtmlSaveOptions for MHTML output.
        // Export all resources (fonts, images) so the result can be viewed offline.
        // -----------------------------------------------------------------
        HtmlSaveOptions mhtmlOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            ExportFontResources = true,               // Embed fonts.
            ExportImagesAsBase64 = true,              // Embed images as Base64.
            ExportCidUrlsForMhtmlResources = true    // Use CID URLs for resources.
        };

        // -----------------------------------------------------------------
        // Step 4: Save the document as MHTML using the configured options.
        // -----------------------------------------------------------------
        epubDoc.Save(mhtmlPath, mhtmlOptions);

        // Verify that the MHTML file was created and contains data.
        if (!File.Exists(mhtmlPath) || new FileInfo(mhtmlPath).Length == 0)
            throw new InvalidOperationException($"MHTML conversion failed; file '{mhtmlPath}' was not created correctly.");

        // Optional: Inform the user that the conversion succeeded.
        Console.WriteLine($"EPUB file '{epubPath}' was successfully converted to MHTML file '{mhtmlPath}'.");
    }

    // Helper method to provide a simple in‑memory image for the sample document.
    // Uses Aspose.Drawing.Common to avoid System.Drawing.
    private static byte[] ImageFromPlaceholder()
    {
        // Create a 100x100 pixel PNG image with a solid color.
        using (var bitmap = new Bitmap(100, 100))
        {
            // Obtain a Graphics object that can draw onto the bitmap.
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightBlue);
            }

            // Save the bitmap to a memory stream in PNG format.
            using (var ms = new MemoryStream())
            {
                bitmap.Save(ms, ImageFormat.Png);
                return ms.ToArray();
            }
        }
    }
}
