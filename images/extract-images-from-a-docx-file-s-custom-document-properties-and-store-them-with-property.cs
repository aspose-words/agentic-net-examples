using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Properties;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample image (100x100 white background with a red rectangle)
        const string sampleImagePath = "sample.png";
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                using (var pen = new Pen(Color.Red, 3))
                {
                    graphics.DrawRectangle(pen, 10, 10, 80, 80);
                }
                bitmap.Save(sampleImagePath, ImageFormat.Png);
            }
        }

        // Read the image bytes and encode them as Base64 for storage in custom document properties
        byte[] imageBytes = File.ReadAllBytes(sampleImagePath);
        string base64Image = Convert.ToBase64String(imageBytes);

        // Create a new empty document
        Document doc = new Document();

        // Add custom document properties that store the image data as Base64 strings
        doc.CustomDocumentProperties.Add("SampleImage", base64Image);
        doc.CustomDocumentProperties.Add("AnotherImage", base64Image);

        // Save the document to a local DOCX file
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // Load the document back (simulating a separate extraction step)
        Document loadedDoc = new Document(docPath);

        // Extract images stored in custom document properties
        int extractedCount = 0;
        foreach (DocumentProperty prop in loadedDoc.CustomDocumentProperties)
        {
            // Only process properties whose value is a Base64 string (image data)
            if (prop.Value is string base64 && !string.IsNullOrEmpty(base64))
            {
                // Decode the Base64 string back to bytes
                byte[] bytes = Convert.FromBase64String(base64);

                // Use the property name as the output file name
                string outputImagePath = $"{prop.Name}.png";

                // Write the image bytes to a file
                File.WriteAllBytes(outputImagePath, bytes);

                // Validate that the file was created
                if (!File.Exists(outputImagePath))
                    throw new InvalidOperationException($"Failed to create image file: {outputImagePath}");

                extractedCount++;
            }
        }

        // Ensure at least one image was extracted
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from custom document properties.");
    }
}
