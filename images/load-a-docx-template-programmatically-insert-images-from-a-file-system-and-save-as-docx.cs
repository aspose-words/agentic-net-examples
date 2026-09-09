using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;          // Aspose.Drawing.Common provides Bitmap, Graphics, Color

namespace AsposeWordsImageInsertExample
{
    public class Program
    {
        public static void Main()
        {
            // Define file names.
            const string templatePath = "Template.docx";
            const string outputPath = "Output.docx";
            const string imagePath1 = "Image1.png";
            const string imagePath2 = "Image2.png";

            // -----------------------------------------------------------------
            // 1. Create a simple DOCX template if it does not already exist.
            // -----------------------------------------------------------------
            if (!File.Exists(templatePath))
            {
                Document templateDoc = new Document();
                DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
                templateBuilder.Writeln("This is a template document.");
                templateDoc.Save(templatePath);
            }

            // ---------------------------------------------------------------
            // 2. Create deterministic sample images using Aspose.Drawing.
            // ---------------------------------------------------------------
            CreateSampleImage(imagePath1, 200, 150, Aspose.Drawing.Color.LightBlue);
            CreateSampleImage(imagePath2, 150, 200, Aspose.Drawing.Color.LightCoral);

            // ---------------------------------------------------------------
            // 3. Load the template document.
            // ---------------------------------------------------------------
            Document doc = new Document(templatePath);
            DocumentBuilder builder = new DocumentBuilder(doc);

            // ---------------------------------------------------------------
            // 4. Insert the images into the document.
            // ---------------------------------------------------------------
            // Ensure the image files exist before insertion.
            if (!File.Exists(imagePath1) || !File.Exists(imagePath2))
                throw new FileNotFoundException("One or more image files were not created.");

            builder.Writeln(); // Add a blank paragraph before images.
            builder.InsertImage(imagePath1);
            builder.Writeln(); // Separate the images with a paragraph break.
            builder.InsertImage(imagePath2);

            // ---------------------------------------------------------------
            // 5. Save the resulting document.
            // ---------------------------------------------------------------
            doc.Save(outputPath, SaveFormat.Docx);

            // ---------------------------------------------------------------
            // 6. Validate that the output file was created.
            // ---------------------------------------------------------------
            if (!File.Exists(outputPath))
                throw new Exception("The output document was not saved correctly.");

            // Optional: clean up generated files (comment out if you want to keep them).
            //File.Delete(templatePath);
            //File.Delete(imagePath1);
            //File.Delete(imagePath2);
        }

        // Helper method to create a PNG image with a solid background color.
        private static void CreateSampleImage(string fileName, int width, int height, Aspose.Drawing.Color backgroundColor)
        {
            // Create a bitmap with the requested dimensions.
            using (Bitmap bitmap = new Bitmap(width, height))
            {
                // Obtain a graphics object to draw on the bitmap.
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    // Fill the entire bitmap with the specified background color.
                    graphics.Clear(backgroundColor);
                }

                // Save the bitmap to the specified file path.
                bitmap.Save(fileName);
            }

            // Verify that the image file was created.
            if (!File.Exists(fileName))
                throw new Exception($"Failed to create image file: {fileName}");
        }
    }
}
