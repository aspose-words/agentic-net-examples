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
        // Create a deterministic sample PNG image.
        const string sampleImagePath = "sample.png";
        CreateSamplePng(sampleImagePath);

        // Create sample Word documents that contain the PNG image.
        string[] docPaths = { "doc1.docx", "doc2.docx" };
        foreach (string docPath in docPaths)
            CreateWordDocWithImage(docPath, sampleImagePath);

        // Apply sepia tone to all extracted PNG images.
        int processedCount = 0;
        foreach (string docPath in docPaths)
        {
            Document doc = new Document(docPath);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapes)
            {
                if (shape.HasImage && shape.ImageData.ImageType == ImageType.Png)
                {
                    // Extract the image bytes.
                    byte[] imageBytes = shape.ImageData.ToByteArray();

                    using (MemoryStream ms = new MemoryStream(imageBytes))
                    {
                        ms.Position = 0;
                        using (Bitmap original = new Bitmap(ms))
                        {
                            using (Bitmap sepia = ApplySepia(original))
                            {
                                string outputPath = $"sepia_{Path.GetFileNameWithoutExtension(docPath)}_{imageIndex}.png";
                                sepia.Save(outputPath);
                                if (!File.Exists(outputPath))
                                    throw new Exception($"Failed to save sepia image: {outputPath}");
                                processedCount++;
                            }
                        }
                    }

                    imageIndex++;
                }
            }
        }

        if (processedCount == 0)
            throw new Exception("No PNG images were found and processed.");
    }

    // Generates a simple PNG image with a red circle.
    private static void CreateSamplePng(string filePath)
    {
        const int width = 200;
        const int height = 200;

        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Color.White);
            using (SolidBrush brush = new SolidBrush(Color.Red))
            {
                g.FillEllipse(brush, 20, 20, 160, 160);
            }
            bitmap.Save(filePath);
        }
    }

    // Creates a Word document and inserts the specified image twice.
    private static void CreateWordDocWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        builder.InsertParagraph();
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Applies a sepia tone using a color matrix.
    private static Bitmap ApplySepia(Bitmap source)
    {
        Bitmap result = new Bitmap(source.Width, source.Height);
        using (Graphics g = Graphics.FromImage(result))
        {
            // Sepia color matrix.
            float[][] matrixElements =
            {
                new float[] { 0.393f, 0.349f, 0.272f, 0, 0 },
                new float[] { 0.769f, 0.686f, 0.534f, 0, 0 },
                new float[] { 0.189f, 0.168f, 0.131f, 0, 0 },
                new float[] { 0,      0,      0,      1, 0 },
                new float[] { 0,      0,      0,      0, 1 }
            };
            ColorMatrix colorMatrix = new ColorMatrix(matrixElements);
            ImageAttributes attributes = new ImageAttributes();
            attributes.SetColorMatrix(colorMatrix);

            g.DrawImage(
                source,
                new Rectangle(0, 0, source.Width, source.Height),
                0,
                0,
                source.Width,
                source.Height,
                GraphicsUnit.Pixel,
                attributes);
        }
        return result;
    }
}
