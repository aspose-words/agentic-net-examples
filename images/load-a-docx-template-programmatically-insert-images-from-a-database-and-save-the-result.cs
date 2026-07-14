using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        const string templatePath = "template.docx";
        const string resultPath = "result.docx";
        const string imagePath1 = "image1.png";
        const string imagePath2 = "image2.png";

        // -------------------------------------------------
        // 1. Create sample images that will act as DB BLOBs
        // -------------------------------------------------
        CreateSampleImage(imagePath1, 100, 100, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage(imagePath2, 100, 100, Aspose.Drawing.Color.LightCoral);

        // -------------------------------------------------
        // 2. Simulate a database that stores image bytes
        // -------------------------------------------------
        List<byte[]> imageBlobs = new List<byte[]>
        {
            File.ReadAllBytes(imagePath1),
            File.ReadAllBytes(imagePath2)
        };

        // -------------------------------------------------
        // 3. Create a simple DOCX template file
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
        templateBuilder.Writeln("Template Document");
        templateBuilder.Writeln("Images will be inserted below:");
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 4. Load the template and insert images from the "DB"
        // -------------------------------------------------
        Document resultDoc = new Document(templatePath);
        DocumentBuilder resultBuilder = new DocumentBuilder(resultDoc);

        foreach (byte[] blob in imageBlobs)
        {
            using (MemoryStream imageStream = new MemoryStream(blob))
            {
                // Ensure the stream is positioned at the beginning
                imageStream.Position = 0;

                // Insert the image inline
                resultBuilder.InsertImage(imageStream);

                // Add a line break after each image
                resultBuilder.Writeln();
            }
        }

        // -------------------------------------------------
        // 5. Save the resulting document
        // -------------------------------------------------
        resultDoc.Save(resultPath);

        // -------------------------------------------------
        // 6. Validate that the output file was created
        // -------------------------------------------------
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The result document was not created.");

        // Cleanup temporary image files (optional)
        File.Delete(imagePath1);
        File.Delete(imagePath2);
        File.Delete(templatePath);
    }

    // Helper method to create a deterministic PNG image using Aspose.Drawing
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color fillColor)
    {
        // Create a bitmap with the requested dimensions
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            // Obtain a graphics object for drawing
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill the background with the specified color
                graphics.Clear(fillColor);
            }

            // Save the bitmap as a PNG file
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}
