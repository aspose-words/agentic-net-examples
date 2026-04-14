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
        // ------------------------------------------------------------
        // Step 1: Create a deterministic sample image (input.png)
        // ------------------------------------------------------------
        const string sampleImagePath = "input.png";
        const int width = 100;
        const int height = 100;

        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (SolidBrush brush = new SolidBrush(Aspose.Drawing.Color.Red))
                {
                    g.FillRectangle(brush, 10, 10, width - 20, height - 20);
                }
            }

            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        if (!File.Exists(sampleImagePath))
            throw new FileNotFoundException("Failed to create the sample image.", sampleImagePath);

        // ------------------------------------------------------------
        // Step 2: Load image bytes and encode as Base64
        // ------------------------------------------------------------
        byte[] imageBytes = File.ReadAllBytes(sampleImagePath);
        string base64Image = Convert.ToBase64String(imageBytes);

        // ------------------------------------------------------------
        // Step 3: Create a DOCX document and add the image data as a custom document property
        // ------------------------------------------------------------
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document containing an image stored in a custom property.");
        doc.CustomDocumentProperties.Add("ImageProp1", base64Image);
        doc.Save(docPath, SaveFormat.Docx);

        if (!File.Exists(docPath))
            throw new FileNotFoundException("Failed to create the sample DOCX.", docPath);

        // ------------------------------------------------------------
        // Step 4: Open the document and extract images from custom document properties
        // ------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        int extractedCount = 0;

        foreach (DocumentProperty prop in loadedDoc.CustomDocumentProperties)
        {
            // Process only string properties that could contain Base64 data
            if (prop.Value is string base64String)
            {
                byte[] data;
                try
                {
                    data = Convert.FromBase64String(base64String);
                }
                catch
                {
                    // Not a valid Base64 string; skip this property
                    continue;
                }

                // Simple PNG signature check (0x89 0x50 0x4E 0x47)
                if (data.Length >= 4 &&
                    data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47)
                {
                    string outputImagePath = $"{prop.Name}.png";

                    using (MemoryStream ms = new MemoryStream(data))
                    {
                        ms.Position = 0; // Ensure stream is at the beginning
                        using (Bitmap extractedBitmap = new Bitmap(ms))
                        {
                            extractedBitmap.Save(outputImagePath, ImageFormat.Png);
                        }
                    }

                    if (!File.Exists(outputImagePath))
                        throw new InvalidOperationException($"Failed to save extracted image for property '{prop.Name}'.");

                    extractedCount++;
                }
            }
        }

        // Validation: at least one image must have been extracted
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from custom document properties.");

        // ------------------------------------------------------------
        // Optional cleanup (comment out if you want to keep the files)
        // ------------------------------------------------------------
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
    }
}
