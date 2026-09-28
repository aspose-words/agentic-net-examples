using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ImageToPdfConverter
{
    public static void Main()
    {
        // Prepare sample images folder
        string imagesFolder = "Images";
        Directory.CreateDirectory(imagesFolder);

        // Define image file paths
        string pngPath = Path.Combine(imagesFolder, "image1.png");
        string jpgPath = Path.Combine(imagesFolder, "image2.jpg");

        // Create a PNG image (200x200, light blue)
        CreateSampleImage(pngPath, 200, 200, Color.LightBlue, ImageFormat.Png);

        // Create a JPEG image (300x150, light coral)
        CreateSampleImage(jpgPath, 300, 150, Color.LightCoral, ImageFormat.Jpeg);

        // List of image files to insert
        string[] imageFiles = new[] { pngPath, jpgPath };

        // Create a new Word document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert each image into the document, adding a page break after each except the last
        for (int i = 0; i < imageFiles.Length; i++)
        {
            string imgPath = imageFiles[i];
            builder.InsertImage(imgPath);
            if (i < imageFiles.Length - 1)
            {
                builder.InsertBreak(BreakType.PageBreak);
            }
        }

        // Save the document as PDF
        string outputPdf = "output.pdf";
        doc.Save(outputPdf, SaveFormat.Pdf);

        // Validate that the PDF was created
        if (!File.Exists(outputPdf))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }

        // Clean up sample images (optional)
        // File.Delete(pngPath);
        // File.Delete(jpgPath);
        // Directory.Delete(imagesFolder);
    }

    private static void CreateSampleImage(string filePath, int width, int height, Color fillColor, ImageFormat format)
    {
        // Create a bitmap with the specified dimensions
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            // Fill the bitmap with the specified color
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(fillColor);
            }

            // Save the bitmap to the given file path using the desired image format
            bitmap.Save(filePath, format);
        }
    }
}
