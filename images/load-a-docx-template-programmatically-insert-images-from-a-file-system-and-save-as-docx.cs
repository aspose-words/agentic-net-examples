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
        // Define deterministic file names.
        const string templatePath = "template.docx";
        const string imagePath = "input.png";
        const string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // 1. Create a sample image using Aspose.Drawing.
        // -----------------------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;

        // Ensure any previous files are removed to avoid stale data.
        if (File.Exists(imagePath)) File.Delete(imagePath);
        if (File.Exists(templatePath)) File.Delete(templatePath);
        if (File.Exists(outputPath)) File.Delete(outputPath);

        // Create bitmap, clear to white, and save.
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                // Optional deterministic drawing – a simple black rectangle.
                using (Pen pen = new Pen(Aspose.Drawing.Color.Black, 3))
                {
                    g.DrawRectangle(pen, 10, 10, imgWidth - 20, imgHeight - 20);
                }
            }

            // Save the image to the file system.
            bitmap.Save(imagePath);
        }

        // Verify that the image file was created.
        if (!File.Exists(imagePath))
            throw new InvalidOperationException($"Image file '{imagePath}' was not created.");

        // -----------------------------------------------------------------
        // 2. Create a DOCX template document.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
        templateBuilder.Writeln("This is a template document.");
        templateDoc.Save(templatePath);

        // Verify that the template file was created.
        if (!File.Exists(templatePath))
            throw new InvalidOperationException($"Template file '{templatePath}' was not created.");

        // -----------------------------------------------------------------
        // 3. Load the template, insert the image, and save the result.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the previously created image using the builder.
        // The InsertImage method appends a Shape with HasImage = true.
        Shape insertedShape = builder.InsertImage(imagePath);

        // Optional: verify that the shape indeed contains an image.
        if (!insertedShape.HasImage)
            throw new InvalidOperationException("Inserted shape does not contain an image.");

        // Save the modified document.
        doc.Save(outputPath);

        // -----------------------------------------------------------------
        // 4. Validate that the output document exists.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Output file '{outputPath}' was not created.");

        // Program completed successfully.
    }
}
