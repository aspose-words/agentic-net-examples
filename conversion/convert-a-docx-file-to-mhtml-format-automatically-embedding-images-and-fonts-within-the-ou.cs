using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX document with some text and an image.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document that will be converted to MHTML.");
        builder.Writeln("The image below is generated programmatically using Aspose.Drawing.");

        // Create a simple bitmap (100x100) filled with blue color.
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
            }

            // Save the bitmap to a memory stream in PNG format.
            using (MemoryStream imageStream = new MemoryStream())
            {
                bitmap.Save(imageStream, ImageFormat.Png);
                imageStream.Position = 0; // Reset for reading.

                // Insert the image into the document.
                builder.InsertImage(imageStream);
            }
        }

        // Save the document as DOCX (required by the task workflow).
        string docxPath = Path.Combine(artifactsDir, "Sample.docx");
        sourceDoc.Save(docxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Load the DOCX document.
        // -----------------------------------------------------------------
        Document doc = new Document(docxPath);

        // -----------------------------------------------------------------
        // 3. Convert to MHTML with embedded images and fonts.
        // -----------------------------------------------------------------
        HtmlSaveOptions mhtmlOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            ExportFontResources = true,               // Embed fonts.
            ExportImagesAsBase64 = false,             // Images will be embedded as MIME parts in MHTML.
            ExportCidUrlsForMhtmlResources = false    // Use default file‑name references.
        };

        string mhtmlPath = Path.Combine(artifactsDir, "Sample.mht");
        doc.Save(mhtmlPath, mhtmlOptions);

        // -----------------------------------------------------------------
        // 4. Validate that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(mhtmlPath) || new FileInfo(mhtmlPath).Length == 0)
            throw new InvalidOperationException("MHTML conversion failed: output file is missing or empty.");

        // The example finishes without waiting for user input.
    }
}
