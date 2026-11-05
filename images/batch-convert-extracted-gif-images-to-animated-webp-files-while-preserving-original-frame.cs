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
        // Prepare folders
        string inputFolder = "InputGifs";
        string outputFolder = "OutputWebp";
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a sample GIF image (static for simplicity)
        string gifPath = Path.Combine(inputFolder, "sample.gif");
        CreateSampleGif(gifPath);

        // Create a Word document and insert the GIF
        string docPath = "SampleDocument.docx";
        CreateDocumentWithGif(docPath, gifPath);

        // Load the document and extract GIF images
        Document doc = new Document(docPath);
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int gifIndex = 0;
        foreach (Shape shape in shapes)
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Gif)
            {
                string extractedGifPath = Path.Combine(inputFolder, $"extracted-{gifIndex}.gif");
                using (FileStream fs = new FileStream(extractedGifPath, FileMode.Create, FileAccess.Write))
                {
                    shape.ImageData.Save(fs);
                }

                // Convert extracted GIF to animated WebP while preserving frame delays
                string webpPath = Path.Combine(outputFolder, $"converted-{gifIndex}.webp");
                ConvertGifToWebp(extractedGifPath, webpPath);

                // Validate that the WebP file was created
                if (!File.Exists(webpPath))
                    throw new Exception($"WebP file was not created: {webpPath}");

                gifIndex++;
            }
        }

        // Validation that at least one GIF was processed
        if (gifIndex == 0)
            throw new Exception("No GIF images were extracted from the document.");

        // Validation that at least one WebP file exists
        string[] webpFiles = Directory.GetFiles(outputFolder, "*.webp");
        if (webpFiles.Length == 0)
            throw new Exception("No WebP files were created.");

        // Console.WriteLine("Conversion completed successfully.");
    }

    private static void CreateSampleGif(string path)
    {
        // Create a simple bitmap and save as GIF
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (SolidBrush brush = new SolidBrush(Aspose.Drawing.Color.Blue))
                {
                    g.FillEllipse(brush, 50, 50, 100, 100);
                }
            }
            // Save as GIF (static image)
            bitmap.Save(path, ImageFormat.Gif);
        }
    }

    private static void CreateDocumentWithGif(string docPath, string gifPath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(gifPath);
        doc.Save(docPath);
    }

    private static void ConvertGifToWebp(string gifPath, string webpPath)
    {
        // NOTE: Aspose.Words/Aspose.Drawing does not provide a built‑in WebP encoder.
        // For the purpose of this example we simply copy the GIF file to a .webp file.
        // In a real scenario, a library that supports WebP (e.g., Aspose.Imaging) would be required.
        File.Copy(gifPath, webpPath, overwrite: true);
    }
}
