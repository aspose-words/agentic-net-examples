using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare deterministic file names
        string imagePath = "cover.png";
        string audioPath = "sample.mp3";
        string docPath = "sample.docx";

        // Create a sample image (100x100 white background with a red rectangle)
        Bitmap bitmap = new Bitmap(100, 100);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        using (Pen pen = new Pen(Color.Red, 5))
        {
            graphics.DrawRectangle(pen, 10, 10, 80, 80);
        }
        bitmap.Save(imagePath);
        graphics.Dispose();
        bitmap.Dispose();

        // Create a dummy audio file (simple MP3 header)
        File.WriteAllBytes(audioPath, new byte[] { 0x49, 0x44, 0x33, 0x03, 0x00, 0x00, 0x00 });

        // Create a new document and insert the image and the audio OLE object
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the sample image (acts as cover art)
        builder.InsertImage(imagePath);

        // Insert the audio file as an OLE object using a FileStream
        using (FileStream audioStream = File.OpenRead(audioPath))
        {
            builder.InsertOleObject(audioStream, "Package", false, null);
        }

        // Save the document
        doc.Save(docPath);

        // Reload the document for extraction
        Document loadedDoc = new Document(docPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;
        int imageIndex = 1;

        foreach (Shape shape in shapes)
        {
            if (shape.HasImage)
            {
                // Save the shape's image to a temporary PNG file
                string tempPngPath = $"temp-{imageIndex}.png";
                shape.ImageData.Save(tempPngPath);

                // Load the temporary PNG and save as JPEG
                using (Image img = Image.FromFile(tempPngPath))
                {
                    string outputPath = $"extracted-{imageIndex}.jpg";
                    img.Save(outputPath, ImageFormat.Jpeg);
                    extractedCount++;
                }

                // Clean up the temporary file
                if (File.Exists(tempPngPath))
                    File.Delete(tempPngPath);

                imageIndex++;
            }
        }

        // Validate that at least one image was extracted
        if (extractedCount == 0)
        {
            throw new InvalidOperationException("No images were extracted from the document.");
        }

        Console.WriteLine($"Extraction complete. {extractedCount} image(s) saved as JPEG.");
    }
}
