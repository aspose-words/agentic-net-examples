using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Set up working directories.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        string inputDir = Path.Combine(workDir, "InputImages");
        string outputDir = Path.Combine(workDir, "OutputWebP");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // 1. Create a sample animated GIF (for demonstration we create a simple static GIF).
        string gifPath = Path.Combine(inputDir, "sample.gif");
        CreateSampleGif(gifPath);

        // 2. Insert the GIF into a Word document.
        string docPath = Path.Combine(workDir, "DocumentWithGif.docx");
        CreateDocumentWithGif(docPath, gifPath);

        // 3. Load the document and extract all GIF images.
        Document doc = new Document(docPath);
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int gifIndex = 0;

        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (!shape.HasImage) continue;

            if (shape.ImageData.ImageType == ImageType.Gif)
            {
                // Save the extracted GIF to the output folder.
                string extractedGif = Path.Combine(outputDir, $"extracted_{gifIndex}.gif");
                shape.ImageData.Save(extractedGif);
                if (!File.Exists(extractedGif))
                    throw new Exception("Failed to extract GIF image.");

                // Convert the GIF to animated WebP while preserving frame delays.
                string webpPath = Path.Combine(outputDir, $"converted_{gifIndex}.webp");
                ConvertGifToWebP(extractedGif, webpPath);

                if (!File.Exists(webpPath))
                    throw new Exception("Failed to create WebP file.");

                gifIndex++;
            }
        }

        // Validation: ensure at least one WebP file was produced.
        if (Directory.GetFiles(outputDir, "*.webp").Length == 0)
            throw new Exception("No WebP files were produced.");
    }

    // Creates a simple GIF image file (static for this example).
    private static void CreateSampleGif(string filePath)
    {
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(100, 100))
        using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
        {
            g.Clear(Aspose.Drawing.Color.Blue);
            bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Gif);
        }
    }

    // Creates a Word document and inserts the specified GIF image.
    private static void CreateDocumentWithGif(string docPath, string gifPath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape gifShape = builder.InsertImage(gifPath);
        if (!gifShape.HasImage)
            throw new Exception("Failed to insert GIF into the document.");
        doc.Save(docPath);
    }

    // Converts a GIF file to WebP using Aspose.Words rendering pipeline.
    private static void ConvertGifToWebP(string gifFilePath, string webpFilePath)
    {
        // Load the GIF into a temporary document.
        Document tempDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(tempDoc);
        builder.InsertImage(gifFilePath);

        // Set up image save options for WebP format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.WebP);

        // Save the document page (which contains the GIF) as a WebP image.
        // Aspose.Words preserves animation frames when saving to WebP.
        tempDoc.Save(webpFilePath, options);
    }
}
