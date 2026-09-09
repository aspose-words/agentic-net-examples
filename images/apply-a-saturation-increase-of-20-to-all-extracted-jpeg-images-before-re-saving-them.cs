using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ImageSaturationExample
{
    public static void Main()
    {
        // Directories for artifacts
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample JPEG image using Aspose.Drawing
        string sampleImagePath = Path.Combine(artifactsDir, "sample.jpg");
        CreateSampleJpeg(sampleImagePath);

        // 2. Create a Word document and insert the JPEG image several times
        string docPath = Path.Combine(artifactsDir, "Original.docx");
        CreateDocumentWithImages(docPath, sampleImagePath);

        // 3. Load the document, increase saturation of each JPEG image by 20%
        string modifiedDocPath = Path.Combine(artifactsDir, "Modified.docx");
        IncreaseJpegSaturation(docPath, modifiedDocPath);

        // 4. Validate that the modified document was saved
        if (!File.Exists(modifiedDocPath))
            throw new InvalidOperationException("The modified document was not saved.");

        Console.WriteLine("Saturation increase completed successfully.");
    }

    // Creates a simple 200x200 JPEG image with a solid rectangle.
    private static void CreateSampleJpeg(string filePath)
    {
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Blue))
                {
                    g.FillRectangle(brush, 25, 25, 150, 150);
                }
            }

            // Save as JPEG
            bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
        }
    }

    // Creates a new document and inserts the same JPEG image three times.
    private static void CreateDocumentWithImages(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three images, each on its own paragraph
        for (int i = 0; i < 3; i++)
        {
            builder.Writeln($"Image #{i + 1}:");
            builder.InsertImage(imagePath);
            builder.Writeln(); // Add spacing
        }

        doc.Save(docPath);
    }

    // Loads the document, processes each JPEG image, increases its saturation by 20%, and saves the result.
    private static void IncreaseJpegSaturation(string inputDocPath, string outputDocPath)
    {
        Document doc = new Document(inputDocPath);

        // Iterate over all Shape nodes that contain images
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only JPEG images
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Extract the image bytes
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Load the image into a Bitmap
            using (MemoryStream srcStream = new MemoryStream(imageBytes))
            using (Aspose.Drawing.Bitmap srcBitmap = new Aspose.Drawing.Bitmap(srcStream))
            {
                // Create a new bitmap to hold the adjusted image
                using (Aspose.Drawing.Bitmap destBitmap = new Aspose.Drawing.Bitmap(srcBitmap.Width, srcBitmap.Height))
                using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(destBitmap))
                {
                    // Build a color matrix that increases saturation by 20%
                    float saturation = 1.2f; // 20% increase
                    float lumR = 0.3086f;
                    float lumG = 0.6094f;
                    float lumB = 0.0820f;

                    float oneMinusSat = 1.0f - saturation;

                    // ColorMatrix expects a jagged array (float[][])
                    float[][] matrixElements = new float[][]
                    {
                        new float[] { oneMinusSat * lumR + saturation, oneMinusSat * lumR,               oneMinusSat * lumR,               0, 0 },
                        new float[] { oneMinusSat * lumG,               oneMinusSat * lumG + saturation, oneMinusSat * lumG,               0, 0 },
                        new float[] { oneMinusSat * lumB,               oneMinusSat * lumB,               oneMinusSat * lumB + saturation, 0, 0 },
                        new float[] { 0,                                 0,                                 0,                                 1, 0 },
                        new float[] { 0,                                 0,                                 0,                                 0, 1 }
                    };

                    Aspose.Drawing.Imaging.ColorMatrix colorMatrix = new Aspose.Drawing.Imaging.ColorMatrix(matrixElements);
                    Aspose.Drawing.Imaging.ImageAttributes imgAttr = new Aspose.Drawing.Imaging.ImageAttributes();
                    imgAttr.SetColorMatrix(colorMatrix, Aspose.Drawing.Imaging.ColorMatrixFlag.Default, Aspose.Drawing.Imaging.ColorAdjustType.Bitmap);

                    // Draw the original bitmap onto the destination using the color matrix
                    graphics.DrawImage(
                        srcBitmap,
                        new Rectangle(0, 0, srcBitmap.Width, srcBitmap.Height),
                        0, 0, srcBitmap.Width, srcBitmap.Height,
                        Aspose.Drawing.GraphicsUnit.Pixel,
                        imgAttr);

                    // Save the adjusted bitmap to a memory stream
                    using (MemoryStream destStream = new MemoryStream())
                    {
                        destBitmap.Save(destStream, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
                        destStream.Position = 0; // Reset before reuse

                        // Replace the image in the shape with the adjusted image
                        shape.ImageData.SetImage(destStream);
                    }
                }
            }
        }

        // Save the modified document
        doc.Save(outputDocPath);
    }
}
