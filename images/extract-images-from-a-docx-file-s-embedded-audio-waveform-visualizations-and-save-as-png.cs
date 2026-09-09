using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing; // Aspose.Drawing.Common namespace

public class ExtractAudioWaveformImages
{
    public static void Main()
    {
        // Define deterministic file names and folders.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        string inputImagePath = Path.Combine(artifactsDir, "waveform.png");
        string docPath = Path.Combine(artifactsDir, "sample.docx");

        // ------------------------------------------------------------
        // 1. Create a sample PNG image that will represent an audio waveform.
        // ------------------------------------------------------------
        const int imgWidth = 400;
        const int imgHeight = 100;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // White background.
                g.Clear(Color.White);

                // Draw a simple waveform-like polyline.
                Pen pen = new Pen(Color.Blue, 2);
                for (int x = 0; x < imgWidth; x += 10)
                {
                    int y = (int)(imgHeight / 2 + 30 * Math.Sin(x * 0.05));
                    g.DrawLine(pen, x, imgHeight / 2, x, y);
                }
                pen.Dispose();
            }

            // Save the image to disk – required before inserting into the document.
            bitmap.Save(inputImagePath);
        }

        // ------------------------------------------------------------
        // 2. Create a DOCX document and insert the waveform image.
        // ------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image as an inline shape.
        builder.InsertImage(inputImagePath);

        // Save the document.
        doc.Save(docPath);

        // ------------------------------------------------------------
        // 3. Load the document and extract all images (including the waveform).
        // ------------------------------------------------------------
        Document loadedDoc = new Document(docPath);

        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine a file name with the proper extension.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string outFile = Path.Combine(artifactsDir, $"extracted_{imageIndex}{extension}");

                // Save the image data to the file system.
                shape.ImageData.Save(outFile);
                Console.WriteLine($"Extracted image saved to: {outFile}");
                imageIndex++;
            }
        }

        // ------------------------------------------------------------
        // 4. Validation – ensure at least one image was extracted.
        // ------------------------------------------------------------
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // The program finishes automatically; no user interaction required.
    }
}
