using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a sample document with some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document with an embedded image.");

        // Generate a simple red square image using Aspose.Drawing.
        using (MemoryStream imageStream = new MemoryStream())
        {
            using (Bitmap bitmap = new Bitmap(100, 100))
            {
                // Fill the bitmap with red color.
                for (int x = 0; x < bitmap.Width; x++)
                {
                    for (int y = 0; y < bitmap.Height; y++)
                    {
                        bitmap.SetPixel(x, y, Color.Red);
                    }
                }

                // Save the bitmap as PNG into the memory stream.
                bitmap.Save(imageStream, ImageFormat.Png);
            }

            // Reset the stream position before inserting.
            imageStream.Position = 0;

            // Insert the image into the document.
            builder.InsertImage(imageStream);
        }

        // Configure HTML save options to embed images as Base64 data URIs.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            ExportImagesAsBase64 = true
        };

        string outputPath = "output.html";
        doc.Save(outputPath, saveOptions);

        // Validate that the HTML file was created and contains data.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The HTML output file was not created.");

        FileInfo info = new FileInfo(outputPath);
        if (info.Length == 0)
            throw new InvalidOperationException("The HTML output file is empty.");

        // Indicate success.
        Console.WriteLine("HTML file with embedded images created successfully.");
    }
}
