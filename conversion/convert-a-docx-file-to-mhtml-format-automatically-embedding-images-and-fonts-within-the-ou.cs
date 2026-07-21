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
        // Define file paths.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);
        string imagePath = Path.Combine(workDir, "sample.png");
        string docxPath = Path.Combine(workDir, "input.docx");
        string mhtmlPath = Path.Combine(workDir, "output.mht");

        // Create a simple PNG image using Aspose.Drawing.
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
            }
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // Create a DOCX document that contains text and the image.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample DOCX content with an embedded image:");
        builder.InsertImage(imagePath);
        sourceDoc.Save(docxPath, SaveFormat.Docx);

        // Load the DOCX document.
        Document doc = new Document(docxPath);

        // Configure save options for MHTML with embedded fonts and images.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            ExportFontResources = true,               // Embed fonts.
            ExportImagesAsBase64 = false,             // Keep images as separate MIME parts (default for MHTML).
            ExportCidUrlsForMhtmlResources = false    // Use file name references (default).
        };

        // Save the document as MHTML.
        doc.Save(mhtmlPath, saveOptions);

        // Validate that the output file was created and is not empty.
        if (!File.Exists(mhtmlPath) || new FileInfo(mhtmlPath).Length == 0)
            throw new InvalidOperationException("MHTML conversion failed: output file was not created or is empty.");

        // Optional cleanup (comment out if you want to inspect the files).
        // File.Delete(imagePath);
        // File.Delete(docxPath);
        // File.Delete(mhtmlPath);
        // Directory.Delete(workDir);
    }
}
