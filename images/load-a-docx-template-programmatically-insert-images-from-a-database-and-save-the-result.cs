using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing; // Aspose.Drawing.Common namespace

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
        // 1. Create a simple DOCX template file.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
        templateBuilder.Writeln("This is a DOCX template.");
        templateBuilder.Writeln("Images will be inserted below:");
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Simulate image data stored in a database.
        //    Create two deterministic PNG images locally.
        // -------------------------------------------------
        CreateSampleImage(imagePath1, Color.Red);
        CreateSampleImage(imagePath2, Color.Green);

        // Read the image files into byte arrays (as if they came from a DB).
        List<byte[]> imageBytesFromDb = new List<byte[]>
        {
            File.ReadAllBytes(imagePath1),
            File.ReadAllBytes(imagePath2)
        };

        // -------------------------------------------------
        // 3. Load the template document.
        // -------------------------------------------------
        Document doc = new Document(templatePath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a blank line before adding images.
        builder.Writeln();

        // -------------------------------------------------
        // 4. Insert each image into the document.
        // -------------------------------------------------
        foreach (byte[] imgBytes in imageBytesFromDb)
        {
            // Insert the image from a byte array.
            builder.InsertImage(imgBytes);
            // Add a line break after each image for readability.
            builder.Writeln();
        }

        // -------------------------------------------------
        // 5. Save the resulting document.
        // -------------------------------------------------
        doc.Save(resultPath);

        // -------------------------------------------------
        // 6. Validate that the output file was created.
        // -------------------------------------------------
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The result document was not created.");

        // Clean up temporary image files (optional).
        File.Delete(imagePath1);
        File.Delete(imagePath2);
        File.Delete(templatePath);
    }

    // Helper method to create a deterministic PNG image.
    private static void CreateSampleImage(string filePath, Color fillColor)
    {
        const int width = 200;
        const int height = 100;

        // Create a bitmap and obtain a graphics object.
        Bitmap bitmap = new Bitmap(width, height);
        Graphics graphics = Graphics.FromImage(bitmap);

        // Fill the background with the specified color.
        graphics.Clear(fillColor);

        // Save the bitmap to a PNG file.
        bitmap.Save(filePath);

        // Release resources.
        graphics.Dispose();
        bitmap.Dispose();
    }
}
