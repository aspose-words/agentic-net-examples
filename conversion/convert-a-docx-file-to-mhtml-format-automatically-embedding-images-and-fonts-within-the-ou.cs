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
        // Prepare a working directory.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // Create a simple PNG image using Aspose.Drawing.
        string imagePath = Path.Combine(workDir, "sample.png");
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            // Obtain a Graphics object from the bitmap.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
            }

            // Save the bitmap as PNG using Aspose.Drawing.Imaging.ImageFormat.
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // Create a sample DOCX document that contains some text and the image.
        string docxPath = Path.Combine(workDir, "input.docx");
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document that will be converted to MHTML.");
        builder.InsertImage(imagePath);
        sourceDoc.Save(docxPath, SaveFormat.Docx);

        // Load the DOCX and convert it to MHTML with embedded resources.
        Document doc = new Document(docxPath);
        HtmlSaveOptions mhtmlOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            ExportFontResources = true,          // embed fonts
            ExportImagesAsBase64 = true,         // embed images
            ExportCidUrlsForMhtmlResources = false // use file name references (default)
        };
        string mhtmlPath = Path.Combine(workDir, "output.mht");
        doc.Save(mhtmlPath, mhtmlOptions);

        // Validate that the MHTML file was created.
        if (!File.Exists(mhtmlPath) || new FileInfo(mhtmlPath).Length == 0)
            throw new InvalidOperationException("MHTML conversion failed: output file was not created or is empty.");

        // Optional cleanup (commented out to allow inspection of generated files).
        // File.Delete(imagePath);
        // File.Delete(docxPath);
    }
}
