using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class PdfToMarkdownConverter
{
    public static void Main()
    {
        // File and folder names
        string pdfPath = "input.pdf";
        string imagePath = "sample.png";
        string markdownPath = "output.md";
        string assetsFolder = "assets";

        // Clean previous run artifacts
        if (File.Exists(pdfPath)) File.Delete(pdfPath);
        if (File.Exists(imagePath)) File.Delete(imagePath);
        if (File.Exists(markdownPath)) File.Delete(markdownPath);
        if (Directory.Exists(assetsFolder)) Directory.Delete(assetsFolder, true);

        // Create a simple PNG image using Aspose.Drawing
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                using (SolidBrush brush = new SolidBrush(Color.Blue))
                {
                    graphics.FillRectangle(brush, new Rectangle(0, 0, 100, 100));
                }
            }
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // Build a sample PDF containing text and the image
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample PDF content with an image:");
        builder.InsertImage(imagePath);
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF for conversion
        Document pdfDoc = new Document(pdfPath);

        // Configure Markdown save options to place images in the "assets" subfolder
        MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
        {
            ImagesFolder = assetsFolder,
            ImagesFolderAlias = assetsFolder
            // ExportImages defaults to true, so no explicit property is needed
        };

        // Perform the conversion
        pdfDoc.Save(markdownPath, mdOptions);

        // Validate that the Markdown file was created
        if (!File.Exists(markdownPath))
            throw new InvalidOperationException("The Markdown file was not created.");

        // Validate that the assets folder exists and contains at least one image
        if (!Directory.Exists(assetsFolder))
            throw new InvalidOperationException("The assets folder was not created.");

        string[] extractedImages = Directory.GetFiles(assetsFolder);
        if (extractedImages.Length == 0)
            throw new InvalidOperationException("No images were extracted to the assets folder.");

        // Optional cleanup (commented out to allow inspection)
        // File.Delete(pdfPath);
        // File.Delete(imagePath);
        // Directory.Delete(assetsFolder, true);
        // File.Delete(markdownPath);
    }
}
