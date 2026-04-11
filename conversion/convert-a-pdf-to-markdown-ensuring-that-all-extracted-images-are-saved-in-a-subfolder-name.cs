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
        // Define working directories
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        string assetsDir = Path.Combine(baseDir, "assets");
        Directory.CreateDirectory(baseDir);
        Directory.CreateDirectory(assetsDir);

        // Create a sample image using Aspose.Drawing
        string imagePath = Path.Combine(baseDir, "sample.png");
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightBlue);
                graphics.DrawEllipse(new Pen(Color.DarkBlue, 5), 20, 20, 160, 160);
            }
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // Build a sample PDF document containing text and the image
        string pdfPath = Path.Combine(baseDir, "sample.pdf");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample PDF document.");
        builder.InsertImage(imagePath);
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF and convert it to Markdown, extracting images to the assets folder
        Document pdfDoc = new Document(pdfPath);
        MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
        {
            ImagesFolder = assetsDir,
            ImagesFolderAlias = "assets"
        };
        string markdownPath = Path.Combine(baseDir, "sample.md");
        pdfDoc.Save(markdownPath, mdOptions);

        // Validation
        if (!File.Exists(markdownPath))
            throw new Exception("Markdown file was not created.");

        if (!Directory.Exists(assetsDir))
            throw new Exception("Assets folder was not created.");

        string[] extractedImages = Directory.GetFiles(assetsDir);
        if (extractedImages.Length == 0)
            throw new Exception("No images were extracted to the assets folder.");

        // Optional: display result paths (commented out to avoid console output)
        // Console.WriteLine($"Markdown saved to: {markdownPath}");
        // Console.WriteLine($"Images saved to: {assetsDir}");
    }
}
