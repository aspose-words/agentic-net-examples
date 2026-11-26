using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Properties;          // Needed for DocumentProperty
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        const string sampleImagePath = "sample.png";
        const string docPath = "sample.docx";

        // -------------------------------------------------
        // 1. Create a deterministic sample PNG image.
        // -------------------------------------------------
        const int width = 200;
        const int height = 100;
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.White);
                // Draw a simple rectangle for visual content
                graphics.DrawRectangle(Aspose.Drawing.Pens.Black, 10, 10, width - 20, height - 20);
            }
            bitmap.Save(sampleImagePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }

        // -------------------------------------------------
        // 2. Load the image bytes and store them as a Base64
        //    string in a custom document property.
        // -------------------------------------------------
        byte[] imageBytes = File.ReadAllBytes(sampleImagePath);
        string base64Image = Convert.ToBase64String(imageBytes);

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with an image stored in a custom property.");

        // Add the custom property named "SampleImage"
        doc.CustomDocumentProperties.Add("SampleImage", base64Image);
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Reload the document and extract images from custom
        //    document properties, saving each with the property name.
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        int extractedCount = 0;

        foreach (DocumentProperty prop in loadedDoc.CustomDocumentProperties)
        {
            // Expect the property value to be a Base64 string representing an image.
            if (prop.Value is string base64 && !string.IsNullOrEmpty(base64))
            {
                try
                {
                    byte[] bytes = Convert.FromBase64String(base64);
                    string outputImagePath = $"{prop.Name}.png";
                    File.WriteAllBytes(outputImagePath, bytes);
                    extractedCount++;

                    // Validate that the file was created.
                    if (!File.Exists(outputImagePath))
                        throw new InvalidOperationException($"Failed to write image file '{outputImagePath}'.");
                }
                catch (FormatException)
                {
                    // Not a valid Base64 string – ignore this property.
                }
            }
        }

        // -------------------------------------------------
        // 4. Validate that at least one image was extracted.
        // -------------------------------------------------
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from custom document properties.");

        // Cleanup temporary files (optional).
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
    }
}
