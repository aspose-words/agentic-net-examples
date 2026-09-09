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
        // Prepare output directories.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outputPath = Path.Combine(artifactsDir, "Dithered.tiff");

        // Create a simple document with a heading and an image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Sample Document");

        // Generate a simple in‑memory image (100×100 blue square) using Aspose.Drawing.
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
            }

            using (MemoryStream imageStream = new MemoryStream())
            {
                // Save the bitmap to the stream as PNG.
                bitmap.Save(imageStream, ImageFormat.Png);
                imageStream.Position = 0;

                // Insert the image into the document from the stream.
                builder.InsertImage(imageStream);
            }
        }

        // Configure TIFF save options to use Floyd‑Steinberg dithering with a high threshold.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt3,
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            ThresholdForFloydSteinbergDithering = (byte)150 // Darken the binary output.
        };

        // Save the document as a TIFF image.
        doc.Save(outputPath, options);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        // Optionally, report success (no interactive prompts required).
        Console.WriteLine("TIFF image saved successfully to: " + outputPath);
    }
}
