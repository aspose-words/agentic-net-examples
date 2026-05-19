using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Properties;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a deterministic sample image.
        const string sampleImagePath = "sample.png";
        CreateSampleImage(sampleImagePath);

        // Step 2: Load image bytes and encode to Base64.
        byte[] imageBytes = File.ReadAllBytes(sampleImagePath);
        string base64Image = Convert.ToBase64String(imageBytes);

        // Step 3: Create a new DOCX document and add the image as a custom property.
        const string docPath = "sample.docx";
        Document doc = new Document();
        doc.CustomDocumentProperties.Add("SampleImage", base64Image);
        doc.Save(docPath);

        // Step 4: Reload the document and extract images from custom properties.
        Document loadedDoc = new Document(docPath);
        int extractedCount = 0;

        foreach (DocumentProperty prop in loadedDoc.CustomDocumentProperties)
        {
            // Process string properties that could contain Base64 image data.
            if (prop.Value is string base64Str && !string.IsNullOrWhiteSpace(base64Str))
            {
                try
                {
                    byte[] decodedBytes = Convert.FromBase64String(base64Str);
                    string outputImagePath = $"{prop.Name}.png";

                    // Write the decoded bytes to a file.
                    using (MemoryStream ms = new MemoryStream(decodedBytes))
                    {
                        ms.Position = 0;
                        using (Image img = Image.FromStream(ms))
                        {
                            img.Save(outputImagePath, ImageFormat.Png);
                        }
                    }

                    if (File.Exists(outputImagePath))
                    {
                        extractedCount++;
                        Console.WriteLine($"Extracted image from property '{prop.Name}' to file '{outputImagePath}'.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to extract image from property '{prop.Name}': {ex.Message}");
                }
            }
        }

        // Validation: ensure at least one image was extracted.
        if (extractedCount == 0)
        {
            throw new InvalidOperationException("No images were extracted from custom document properties.");
        }

        // Cleanup: optional removal of temporary files (commented out to keep results).
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
    }

    private static void CreateSampleImage(string filePath)
    {
        const int width = 200;
        const int height = 200;

        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                using (Pen pen = new Pen(Color.Blue, 5))
                {
                    graphics.DrawRectangle(pen, 20, 20, width - 40, height - 40);
                }
            }

            bitmap.Save(filePath, ImageFormat.Png);
        }

        if (!File.Exists(filePath))
        {
            throw new InvalidOperationException($"Failed to create sample image at '{filePath}'.");
        }
    }
}
