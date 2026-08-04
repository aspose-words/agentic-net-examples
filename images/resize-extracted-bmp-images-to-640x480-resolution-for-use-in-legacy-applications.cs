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
        // Create a sample BMP image.
        const string sampleBmpPath = "sample.bmp";
        using (Bitmap bmp = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.LightBlue);
                // Draw a simple rectangle.
                g.DrawRectangle(new Pen(Color.DarkBlue, 5), 20, 20, 160, 160);
            }
            bmp.Save(sampleBmpPath);
        }

        // Create a Word document and insert the BMP image.
        const string docPath = "document.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleBmpPath);
        doc.Save(docPath);

        // Load the document and extract images.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Extract image bytes.
            byte[] imageBytes = shape.ImageData.ImageBytes;
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                ms.Position = 0;
                using (Bitmap original = new Bitmap(ms))
                {
                    // Resize to 640x480.
                    using (Bitmap resized = new Bitmap(640, 480))
                    {
                        using (Graphics g = Graphics.FromImage(resized))
                        {
                            g.Clear(Color.White);
                            g.DrawImage(original, 0, 0, 640, 480);
                        }

                        string resizedPath = $"resized-{imageIndex}.bmp";
                        resized.Save(resizedPath);
                        imageIndex++;
                    }
                }
            }
        }

        // Validate that at least one resized image was created.
        string[] resizedFiles = Directory.GetFiles(Directory.GetCurrentDirectory(), "resized-*.bmp");
        if (resizedFiles.Length == 0)
            throw new InvalidOperationException("No resized BMP images were created.");

        // Cleanup (optional): delete temporary files.
        // File.Delete(sampleBmpPath);
        // File.Delete(docPath);
    }
}
