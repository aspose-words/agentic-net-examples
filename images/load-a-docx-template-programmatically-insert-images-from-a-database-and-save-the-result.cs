using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing; // Aspose.Drawing.Common provides Bitmap, Graphics, Color

public class Program
{
    public static void Main()
    {
        // Paths for files used in the example
        const string templatePath = "template.docx";
        const string resultPath = "result.docx";
        const string imagePath1 = "sample1.png";
        const string imagePath2 = "sample2.png";

        // ------------------------------------------------------------
        // 1. Create sample images that will act as data retrieved from a DB.
        // ------------------------------------------------------------
        CreateSampleImage(imagePath1, 200, 100, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage(imagePath2, 150, 150, Aspose.Drawing.Color.LightCoral);

        // Simulate image BLOBs stored in a database as byte arrays.
        List<byte[]> imageBlobs = new List<byte[]>
        {
            File.ReadAllBytes(imagePath1),
            File.ReadAllBytes(imagePath2)
        };

        // ------------------------------------------------------------
        // 2. Create a simple DOCX template that will be loaded later.
        // ------------------------------------------------------------
        CreateTemplateDocument(templatePath);

        // ------------------------------------------------------------
        // 3. Load the template document.
        // ------------------------------------------------------------
        Document doc = new Document(templatePath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder to the end of the document to insert images.
        builder.MoveToDocumentEnd();

        // ------------------------------------------------------------
        // 4. Insert each image from the simulated DB into the document.
        // ------------------------------------------------------------
        foreach (byte[] imageBytes in imageBlobs)
        {
            // Use a MemoryStream for the image bytes.
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                ms.Position = 0; // Ensure the stream is at the beginning.
                // Insert the image inline.
                builder.InsertImage(ms);
                // Add a line break after each image for readability.
                builder.InsertBreak(BreakType.LineBreak);
            }
        }

        // ------------------------------------------------------------
        // 5. Save the resulting document.
        // ------------------------------------------------------------
        doc.Save(resultPath, SaveFormat.Docx);

        // ------------------------------------------------------------
        // 6. Validate that the output file was created.
        // ------------------------------------------------------------
        if (!File.Exists(resultPath))
            throw new InvalidOperationException($"The output file '{resultPath}' was not created.");

        // Clean up temporary files (optional).
        // File.Delete(imagePath1);
        // File.Delete(imagePath2);
        // File.Delete(templatePath);
    }

    // Helper method to create a deterministic PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string fileName, int width, int height, Aspose.Drawing.Color backgroundColor)
    {
        // Create a bitmap with the desired dimensions.
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            // Obtain a graphics object to draw on the bitmap.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill the background with a solid color.
                graphics.Clear(backgroundColor);
            }

            // Save the bitmap to a PNG file.
            bitmap.Save(fileName);
        }
    }

    // Helper method to create a minimal DOCX template.
    private static void CreateTemplateDocument(string fileName)
    {
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("This is a template document.");
        builder.Writeln("Images will be inserted below:");
        // Add an empty paragraph to separate from inserted images.
        builder.InsertParagraph();

        templateDoc.Save(fileName, SaveFormat.Docx);
    }
}
