using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare deterministic folders
        string inputImagePath = "input.png";
        string documentPath = "sample.docx";
        string archiveFolder = "SecureArchive";

        // Ensure the archive folder exists
        Directory.CreateDirectory(archiveFolder);

        // -------------------------------------------------
        // Step 1: Create a sample PNG image using Aspose.Drawing
        // -------------------------------------------------
        const int width = 200;
        const int height = 200;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill background with white
                g.Clear(Color.White);
                // Draw a simple red rectangle
                g.FillRectangle(new SolidBrush(Color.Red), 50, 50, 100, 100);
            }
            // Save the PNG image to a deterministic file
            bitmap.Save(inputImagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // Step 2: Create a Word document and insert the PNG image
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        // Save the document (optional, just to demonstrate lifecycle)
        doc.Save(documentPath);

        // -------------------------------------------------
        // Step 3: Extract PNG images, convert each to grayscale BMP, and save to the secure folder
        // -------------------------------------------------
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Save the shape image to a memory stream
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0; // Reset before reading

                // Load the PNG into a bitmap
                using (Bitmap pngBitmap = new Bitmap(imageStream))
                {
                    // Create a new bitmap for the grayscale version
                    using (Bitmap grayBitmap = new Bitmap(pngBitmap.Width, pngBitmap.Height))
                    {
                        // Convert each pixel to grayscale
                        for (int y = 0; y < pngBitmap.Height; y++)
                        {
                            for (int x = 0; x < pngBitmap.Width; x++)
                            {
                                Color originalColor = pngBitmap.GetPixel(x, y);
                                // Compute luminance using standard Rec. 601 coefficients
                                int luminance = (int)(0.299 * originalColor.R + 0.587 * originalColor.G + 0.114 * originalColor.B);
                                Color grayColor = Color.FromArgb(luminance, luminance, luminance);
                                grayBitmap.SetPixel(x, y, grayColor);
                            }
                        }

                        // Save the grayscale bitmap as BMP in the secure archive folder
                        string outputPath = Path.Combine(archiveFolder, $"image-{imageIndex}.bmp");
                        grayBitmap.Save(outputPath, ImageFormat.Bmp);
                        imageIndex++;
                    }
                }
            }
        }

        // -------------------------------------------------
        // Validation: Ensure at least one BMP file was created
        // -------------------------------------------------
        string[] bmpFiles = Directory.GetFiles(archiveFolder, "*.bmp");
        if (bmpFiles.Length == 0)
        {
            throw new InvalidOperationException("No BMP files were created in the secure archive folder.");
        }

        // Optional: Clean up intermediate files (comment out if you need to inspect them)
        // File.Delete(inputImagePath);
        // File.Delete(documentPath);
    }
}
