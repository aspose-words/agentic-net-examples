using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document with text.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Font.Name = "Arial";
        builder.Writeln("Sample text with embedded font.");

        // Create a simple bitmap image using Aspose.Drawing.
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightBlue);
                using (Pen pen = new Pen(Color.Red, 2))
                {
                    graphics.DrawEllipse(pen, 10, 10, 80, 80);
                }
            }

            // Save the bitmap to a memory stream as PNG.
            using (MemoryStream imageStream = new MemoryStream())
            {
                bitmap.Save(imageStream, ImageFormat.Png);
                imageStream.Position = 0;

                // Insert the image into the document.
                builder.InsertImage(imageStream);
            }
        }

        // Save the sample document as DOCX (input file).
        const string inputPath = "sample.docx";
        sampleDoc.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX document for conversion.
        Document doc = new Document(inputPath);

        // Save the document as MHTML. Aspose.Words embeds images and fonts by default.
        const string outputPath = "output.mhtml";
        doc.Save(outputPath, SaveFormat.Mhtml);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Expected output MHTML was not created.");
        }

        // Optional cleanup of the sample input file.
        // File.Delete(inputPath);
    }
}
