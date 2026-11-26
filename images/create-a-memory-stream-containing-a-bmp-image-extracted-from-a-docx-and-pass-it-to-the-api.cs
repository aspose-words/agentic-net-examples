using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Create a deterministic BMP image.
        const int width = 100;
        const int height = 100;
        string bmpPath = "sample.bmp";

        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightBlue);
                using (SolidBrush brush = new SolidBrush(Color.Red))
                {
                    graphics.FillRectangle(brush, 10, 10, 80, 80);
                }
            }
            bitmap.Save(bmpPath, ImageFormat.Bmp);
        }

        // Create a DOCX and insert the BMP image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(bmpPath);
        string docPath = "sample.docx";
        doc.Save(docPath);

        // Load the document (optional, demonstrates typical workflow).
        Document loadedDoc = new Document(docPath);

        // Extract the first image from the document into a MemoryStream as BMP.
        MemoryStream imageStream = null;
        foreach (Shape shape in loadedDoc.GetChildNodes(NodeType.Shape, true))
        {
            if (shape.HasImage)
            {
                imageStream = new MemoryStream();
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0; // Reset for downstream consumption.
                break;
            }
        }

        if (imageStream == null || imageStream.Length == 0)
            throw new InvalidOperationException("No image was extracted from the document.");

        // Example API usage: display the size of the extracted image stream.
        Console.WriteLine($"Extracted image stream length: {imageStream.Length} bytes");

        // Clean up temporary files.
        if (File.Exists(bmpPath))
            File.Delete(bmpPath);
        if (File.Exists(docPath))
            File.Delete(docPath);
    }
}
