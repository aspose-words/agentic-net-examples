using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

namespace AsposeWordsImageInsertExample
{
    public class Program
    {
        public static void Main()
        {
            // Paths for temporary files
            const string imagePath1 = "image1.png";
            const string imagePath2 = "image2.png";
            const string templatePath = "template.docx";
            const string resultPath = "result.docx";

            // -------------------------------------------------
            // 1. Create sample images that will act as DB BLOBs
            // -------------------------------------------------
            CreateSampleImage(imagePath1, Aspose.Drawing.Color.LightBlue);
            CreateSampleImage(imagePath2, Aspose.Drawing.Color.LightCoral);

            // Read the images into byte arrays to simulate database storage
            var imagesFromDatabase = new List<byte[]>
            {
                File.ReadAllBytes(imagePath1),
                File.ReadAllBytes(imagePath2)
            };

            // -------------------------------------------------
            // 2. Create a simple DOCX template
            // -------------------------------------------------
            var templateDoc = new Document();
            var templateBuilder = new DocumentBuilder(templateDoc);
            templateBuilder.Writeln("Template Document");
            templateBuilder.Writeln("Below are images inserted from the database:");
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 3. Load the template and insert images
            // -------------------------------------------------
            var doc = new Document(templatePath);
            var builder = new DocumentBuilder(doc);

            // Move the cursor to the end of the document (after the placeholder text)
            builder.MoveToDocumentEnd();

            foreach (var imageBytes in imagesFromDatabase)
            {
                // Insert the image from a byte array
                builder.InsertImage(imageBytes);
                // Add a line break after each image for readability
                builder.Writeln();
            }

            // -------------------------------------------------
            // 4. Save the resulting document
            // -------------------------------------------------
            doc.Save(resultPath);

            // -------------------------------------------------
            // 5. Validate that the output file was created
            // -------------------------------------------------
            if (!File.Exists(resultPath))
                throw new InvalidOperationException($"The result document '{resultPath}' was not created.");

            // Clean up temporary image files (optional)
            // File.Delete(imagePath1);
            // File.Delete(imagePath2);
        }

        private static void CreateSampleImage(string filePath, Aspose.Drawing.Color backgroundColor)
        {
            // Create a 100x100 bitmap and fill it with the specified background color
            using (var bitmap = new Aspose.Drawing.Bitmap(100, 100))
            {
                using (var graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
                {
                    graphics.Clear(backgroundColor);
                }

                // Save the bitmap to a PNG file
                bitmap.Save(filePath);
            }
        }
    }
}
