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
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample DOCX document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose.Words!");

        // Create a simple bitmap image using Aspose.Drawing.
        string imagePath = Path.Combine(outputDir, "sample.png");
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
            }
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // Insert the image into the document.
        builder.InsertImage(imagePath);

        // Save the document as DOCX (required by the task rules).
        string docxPath = Path.Combine(outputDir, "sample.docx");
        doc.Save(docxPath, SaveFormat.Docx);

        // Load the DOCX document for conversion.
        Document loadDoc = new Document(docxPath);

        // Configure save options to embed images and fonts into MHTML.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            ExportImagesAsBase64 = true,      // Embed images as Base64.
            ExportFontResources = true,       // Export font resources.
            ExportFontsAsBase64 = true,       // Embed fonts as Base64.
            PrettyFormat = true               // Optional: make output more readable.
        };

        // Save the document as MHTML.
        string mhtmlPath = Path.Combine(outputDir, "sample.mht");
        loadDoc.Save(mhtmlPath, saveOptions);

        // Validate that the MHTML file was created and contains data.
        if (!File.Exists(mhtmlPath) || new FileInfo(mhtmlPath).Length == 0)
        {
            throw new InvalidOperationException("MHTML conversion failed: output file is missing or empty.");
        }

        Console.WriteLine($"Conversion successful. MHTML file saved to: {mhtmlPath}");
    }
}
