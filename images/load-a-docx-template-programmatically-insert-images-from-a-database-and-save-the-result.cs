using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing; // Aspose.Drawing.Common provides Bitmap, Graphics, Color

public class Program
{
    // Entry point
    public static void Main()
    {
        // Paths for temporary files
        const string templatePath = "template.docx";
        const string resultPath = "result.docx";

        // Step 1: Create sample images that simulate database BLOBs
        List<byte[]> imageBlobs = CreateSampleImageBlobs();

        // Step 2: Create a simple DOCX template
        CreateTemplateDocument(templatePath);

        // Step 3: Load the template
        Document doc = new Document(templatePath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move cursor to the end of the document
        builder.MoveToDocumentEnd();

        // Step 4: Insert each image from the "database"
        foreach (byte[] imageBytes in imageBlobs)
        {
            // Insert image from byte array (inline)
            builder.InsertImage(imageBytes);
            // Add a line break after each image for readability
            builder.InsertBreak(BreakType.LineBreak);
        }

        // Step 5: Save the resulting document
        doc.Save(resultPath);

        // Validation: ensure the output file was created
        if (!File.Exists(resultPath))
            throw new InvalidOperationException($"Failed to create output file: {resultPath}");
    }

    // Generates deterministic sample images and returns them as byte arrays
    private static List<byte[]> CreateSampleImageBlobs()
    {
        var blobs = new List<byte[]>();

        // First image: red square
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                g.FillRectangle(new SolidBrush(Color.Red), 20, 20, 160, 160);
            }

            using (MemoryStream ms = new MemoryStream())
            {
                bitmap.Save(ms, Aspose.Drawing.Imaging.ImageFormat.Png);
                ms.Position = 0; // Reset before reading
                blobs.Add(ms.ToArray());
            }
        }

        // Second image: green circle
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                g.FillEllipse(new SolidBrush(Color.Green), 20, 20, 160, 160);
            }

            using (MemoryStream ms = new MemoryStream())
            {
                bitmap.Save(ms, Aspose.Drawing.Imaging.ImageFormat.Png);
                ms.Position = 0;
                blobs.Add(ms.ToArray());
            }
        }

        // Ensure at least one image was created
        if (blobs.Count == 0)
            throw new InvalidOperationException("No sample images were generated.");

        return blobs;
    }

    // Creates a minimal DOCX file that will serve as the template
    private static void CreateTemplateDocument(string path)
    {
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Template Document");
        builder.Writeln("Images will be inserted below:");
        builder.InsertBreak(BreakType.PageBreak);

        template.Save(path);

        // Validate template creation
        if (!File.Exists(path))
            throw new InvalidOperationException($"Failed to create template file: {path}");
    }
}
