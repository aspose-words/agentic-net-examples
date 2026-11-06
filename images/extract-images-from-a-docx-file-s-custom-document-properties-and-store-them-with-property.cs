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
        // Paths for temporary files
        const string imagePath = "sample.png";
        const string docPath = "sample.docx";

        // -------------------------------------------------
        // Create a deterministic sample image (100x100 white)
        // -------------------------------------------------
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(100, 100);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        // Save the image to a file (required by the workflow)
        bitmap.Save(imagePath, ImageFormat.Png);
        graphics.Dispose();
        bitmap.Dispose();

        // Read the image bytes to store in a custom document property
        byte[] imageBytes = File.ReadAllBytes(imagePath);
        // Store as Base64 string because CustomDocumentProperties does not accept raw byte[]
        string base64Image = Convert.ToBase64String(imageBytes);

        // -------------------------------------------------
        // Create a new DOCX document and add the image bytes
        // as a custom document property
        // -------------------------------------------------
        Document doc = new Document();
        CustomDocumentProperties props = doc.CustomDocumentProperties;
        // Property name will be used as the output file name later
        props.Add("Image1", base64Image);
        // Save the document
        doc.Save(docPath);

        // -------------------------------------------------
        // Load the document and extract images from custom properties
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        CustomDocumentProperties loadedProps = loadedDoc.CustomDocumentProperties;

        int extractedCount = 0;
        foreach (DocumentProperty prop in loadedProps)
        {
            // Only process properties that hold a Base64 string (image data)
            if (prop.Value is string s && !string.IsNullOrEmpty(s))
            {
                try
                {
                    byte[] bytes = Convert.FromBase64String(s);
                    if (bytes.Length > 0)
                    {
                        string outputImagePath = $"{prop.Name}.png";
                        File.WriteAllBytes(outputImagePath, bytes);
                        extractedCount++;
                        Console.WriteLine($"Extracted image from property '{prop.Name}' to file '{outputImagePath}'.");
                    }
                }
                catch (FormatException)
                {
                    // Not a Base64 string – ignore
                }
            }
        }

        // Validate that at least one image was extracted
        if (extractedCount == 0)
            throw new InvalidOperationException("No image data found in custom document properties.");

        // Clean up temporary files (optional)
        // File.Delete(imagePath);
        // File.Delete(docPath);
    }
}
