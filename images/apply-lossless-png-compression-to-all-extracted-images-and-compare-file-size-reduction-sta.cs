using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare directories
        string workDir = Directory.GetCurrentDirectory();

        // 1. Create a sample PNG image using Aspose.Drawing
        string sampleImagePath = Path.Combine(workDir, "sample.png");
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.LightBlue);
                // Draw a simple rectangle
                g.DrawRectangle(new Aspose.Drawing.Pen(Aspose.Drawing.Color.DarkBlue, 5), 20, 20, 160, 160);
            }
            bitmap.Save(sampleImagePath); // Extension .png ensures PNG format
        }

        // 2. Create a Word document and insert the sample image multiple times
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document with images:");
        builder.InsertImage(sampleImagePath);
        builder.InsertParagraph();
        builder.InsertImage(sampleImagePath);
        string docPath = Path.Combine(workDir, "Original.docx");
        doc.Save(docPath);

        // 3. Load the document and extract all images
        Document loadedDoc = new Document(docPath);
        var shapeImages = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                   .Cast<Shape>()
                                   .Where(s => s.HasImage)
                                   .ToList();

        if (!shapeImages.Any())
            throw new InvalidOperationException("No images were found in the document.");

        int index = 0;
        foreach (var shape in shapeImages)
        {
            // Determine original file extension
            string ext = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
            string originalFile = Path.Combine(workDir, $"extracted_original_{index}{ext}");
            shape.ImageData.Save(originalFile);
            long originalSize = new FileInfo(originalFile).Length;

            // 4. Apply lossless PNG compression (re‑encode as PNG)
            string compressedFile = Path.Combine(workDir, $"compressed_{index}.png");
            using (MemoryStream ms = new MemoryStream())
            {
                shape.ImageData.Save(ms);
                ms.Position = 0;
                using (Bitmap bmp = new Bitmap(ms))
                {
                    // Saving with .png extension writes a losslessly compressed PNG
                    bmp.Save(compressedFile);
                }
            }
            long compressedSize = new FileInfo(compressedFile).Length;

            // 5. Output size reduction statistics
            long reduction = originalSize - compressedSize;
            Console.WriteLine($"Image {index}: Original = {originalSize} bytes, Compressed = {compressedSize} bytes, Reduction = {reduction} bytes.");

            index++;
        }
    }
}
