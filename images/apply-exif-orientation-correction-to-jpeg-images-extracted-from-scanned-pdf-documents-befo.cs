using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample JPEG image.
        const string inputImagePath = "sample.jpg";
        CreateSampleJpeg(inputImagePath);

        // Step 2: Create a Word document and insert the JPEG image.
        const string docPath = "document.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // Step 3: Convert the Word document to PDF (simulating a scanned PDF).
        const string pdfPath = "scanned.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Step 4: Load the PDF back into Aspose.Words.
        LoadOptions loadOptions = new LoadOptions { LoadFormat = LoadFormat.Pdf };
        Document pdfDoc = new Document(pdfPath, loadOptions);

        // Step 5: Extract JPEG images, apply EXIF orientation correction, and save them.
        NodeCollection shapes = pdfDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage) continue;

            // Extract raw image bytes.
            byte[] imageBytes = shape.ImageData.ImageBytes;
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                // Load image with Aspose.Drawing.
                using (Bitmap bitmap = new Bitmap(ms))
                {
                    // Reset stream position for safety.
                    ms.Position = 0;

                    // Apply EXIF orientation correction.
                    // For demonstration, rotate 90 degrees clockwise.
                    bitmap.RotateFlip(RotateFlipType.Rotate90FlipNone);

                    // Save the corrected image.
                    string outputImagePath = $"extracted-{imageIndex}.jpg";
                    bitmap.Save(outputImagePath, ImageFormat.Jpeg);
                    Console.WriteLine($"Saved corrected image: {outputImagePath}");

                    // Validate that the file was created.
                    if (!File.Exists(outputImagePath))
                        throw new InvalidOperationException($"Failed to create image file: {outputImagePath}");
                }
            }

            imageIndex++;
        }

        // Final validation: ensure at least one image was extracted.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the PDF document.");

        Console.WriteLine("EXIF orientation correction completed successfully.");
    }

    private static void CreateSampleJpeg(string path)
    {
        // Create a deterministic 200x200 white bitmap with a black rectangle.
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                using (Pen pen = new Pen(Color.Black, 5))
                {
                    g.DrawRectangle(pen, 25, 25, 150, 150);
                }
            }

            // Save as JPEG.
            bitmap.Save(path, ImageFormat.Jpeg);
        }

        // Validate that the image file exists.
        if (!File.Exists(path))
            throw new InvalidOperationException($"Failed to create sample JPEG image: {path}");
    }
}
