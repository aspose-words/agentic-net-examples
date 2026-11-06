using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample PNG image.
        const int imgWidth = 200;
        const int imgHeight = 200;
        string sampleImagePath = "sample.png";

        Bitmap bitmap = new Bitmap(imgWidth, imgHeight);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        using (Pen pen = new Pen(Color.Blue, 5))
        {
            graphics.DrawRectangle(pen, 10, 10, imgWidth - 20, imgHeight - 20);
        }
        bitmap.Save(sampleImagePath, ImageFormat.Png);
        graphics.Dispose();
        bitmap.Dispose();

        // Create a Word document and insert the sample image twice.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        builder.Writeln();
        builder.InsertImage(sampleImagePath);
        string docPath = "sample.docx";
        doc.Save(docPath);

        // Load the document and extract all images.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        List<string> originalFiles = new List<string>();
        List<string> compressedFiles = new List<string>();
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Save the original extracted image.
            string originalPath = $"extracted-{imageIndex}.png";
            using (MemoryStream ms = new MemoryStream())
            {
                shape.ImageData.Save(ms);
                ms.Position = 0;
                File.WriteAllBytes(originalPath, ms.ToArray());
            }
            originalFiles.Add(originalPath);

            // Load the image into a bitmap and re‑save it (lossless PNG compression).
            string compressedPath = $"compressed-{imageIndex}.png";
            using (MemoryStream ms = new MemoryStream(File.ReadAllBytes(originalPath)))
            {
                ms.Position = 0;
                using (Bitmap bmp = new Bitmap(ms))
                {
                    bmp.Save(compressedPath, ImageFormat.Png);
                }
            }
            compressedFiles.Add(compressedPath);
            imageIndex++;
        }

        if (originalFiles.Count == 0)
            throw new Exception("No images were extracted from the document.");

        // Compare file sizes and output statistics.
        for (int i = 0; i < originalFiles.Count; i++)
        {
            long originalSize = new FileInfo(originalFiles[i]).Length;
            long compressedSize = new FileInfo(compressedFiles[i]).Length;
            long reduction = originalSize - compressedSize;
            Console.WriteLine($"Image {i}: Original = {originalSize} bytes, Compressed = {compressedSize} bytes, Reduction = {reduction} bytes");
        }
    }
}
