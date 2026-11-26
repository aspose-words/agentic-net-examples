using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Define base, input and output directories.
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string outputFolder = Path.Combine(baseDir, "ExtractedImages");

        // Ensure a clean environment.
        if (Directory.Exists(inputFolder)) Directory.Delete(inputFolder, true);
        if (Directory.Exists(outputFolder)) Directory.Delete(outputFolder, true);
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create deterministic sample images (PNG and JPEG).
        string pngPath = Path.Combine(baseDir, "sample.png");
        string jpegPath = Path.Combine(baseDir, "sample.jpg");

        CreateSampleImage(pngPath, 200, 100, Aspose.Drawing.Color.LightBlue, Aspose.Drawing.Color.Black, "PNG");
        CreateSampleImage(jpegPath, 150, 150, Aspose.Drawing.Color.LightCoral, Aspose.Drawing.Color.White, "JPG");

        // Build sample DOCX files that contain the images.
        for (int i = 1; i <= 2; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Document {i}");
            builder.InsertImage(pngPath);
            builder.InsertParagraph();
            builder.InsertImage(jpegPath);
            string docPath = Path.Combine(inputFolder, $"Doc{i}.docx");
            doc.Save(docPath);
        }

        // Process each DOCX file: extract images and place them into format‑specific subfolders.
        foreach (string docFile in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(docFile);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapes)
            {
                if (!shape.HasImage) continue;

                ImageType imgType = shape.ImageData.ImageType;
                string formatFolderName = imgType.ToString().ToLower(); // e.g., "png", "jpeg"
                string formatFolderPath = Path.Combine(outputFolder, formatFolderName);
                Directory.CreateDirectory(formatFolderPath);

                string extension = GetExtensionFromImageType(imgType);
                string outputFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_img{imageIndex}{extension}";
                string outputPath = Path.Combine(formatFolderPath, outputFileName);

                shape.ImageData.Save(outputPath);
                imageIndex++;
            }

            if (imageIndex == 0)
                throw new InvalidOperationException($"No images were extracted from document: {docFile}");
        }

        // Validate that at least one image was saved.
        bool anyImageSaved = false;
        foreach (string subDir in Directory.GetDirectories(outputFolder))
        {
            if (Directory.GetFiles(subDir).Length > 0)
            {
                anyImageSaved = true;
                break;
            }
        }

        if (!anyImageSaved)
            throw new InvalidOperationException("Image extraction completed but no images were saved.");

        // Clean up the temporary sample image files.
        File.Delete(pngPath);
        File.Delete(jpegPath);
    }

    // Creates a simple bitmap with background colour and a text label.
    private static void CreateSampleImage(string filePath, int width, int height,
        Aspose.Drawing.Color backColor, Aspose.Drawing.Color textColor, string label)
    {
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                graphics.Clear(backColor);
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 20))
                using (Aspose.Drawing.Brush brush = new Aspose.Drawing.SolidBrush(textColor))
                {
                    graphics.DrawString(label, font, brush, new Aspose.Drawing.PointF(10, 40));
                }
            }
            bitmap.Save(filePath);
        }
    }

    // Maps Aspose.Words.ImageType values to common file extensions.
    private static string GetExtensionFromImageType(ImageType type)
    {
        switch (type)
        {
            case ImageType.Jpeg:
                return ".jpg";
            case ImageType.Png:
                return ".png";
            case ImageType.Gif:
                return ".gif";
            case ImageType.Bmp:
                return ".bmp";
            case ImageType.Emf:
                return ".emf";
            case ImageType.Wmf:
                return ".wmf";
            // Aspose.Words does not define a TIFF enum value; treat unknown types generically.
            default:
                return ".img";
        }
    }
}
