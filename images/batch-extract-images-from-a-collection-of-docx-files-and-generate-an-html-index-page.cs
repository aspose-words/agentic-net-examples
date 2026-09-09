using System;
using System.Collections.Generic;
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
        // Root folder for all temporary data.
        string rootFolder = Path.Combine(Directory.GetCurrentDirectory(), "BatchImageExtraction");
        string inputFolder = Path.Combine(rootFolder, "InputDocs");
        string imagesFolder = Path.Combine(rootFolder, "ExtractedImages");
        string htmlIndexPath = Path.Combine(rootFolder, "index.html");

        // Ensure clean environment.
        if (Directory.Exists(rootFolder))
            Directory.Delete(rootFolder, true);
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(imagesFolder);

        // -------------------------------------------------
        // 1. Create a deterministic sample image (input.png).
        // -------------------------------------------------
        string sampleImagePath = Path.Combine(rootFolder, "input.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // -------------------------------------------------
        // 2. Create a few sample DOCX files that contain the image.
        // -------------------------------------------------
        const int docCount = 3;
        for (int i = 1; i <= docCount; i++)
        {
            string docPath = Path.Combine(inputFolder, $"Document{i}.docx");
            CreateDocumentWithImage(docPath, sampleImagePath);
        }

        // -------------------------------------------------
        // 3. Batch process all DOCX files: extract images.
        // -------------------------------------------------
        List<string> extractedImagePaths = new List<string>();
        foreach (string docFile in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(docFile);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_Image{imageIndex}{extension}";
                string imageFullPath = Path.Combine(imagesFolder, imageFileName);

                shape.ImageData.Save(imageFullPath);
                extractedImagePaths.Add(imageFullPath);
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted.
        if (extractedImagePaths.Count == 0)
            throw new InvalidOperationException("No images were extracted from the DOCX files.");

        // -------------------------------------------------
        // 4. Generate a simple HTML index page linking to the images.
        // -------------------------------------------------
        using (StreamWriter writer = new StreamWriter(htmlIndexPath, false))
        {
            writer.WriteLine("<!DOCTYPE html>");
            writer.WriteLine("<html><head><meta charset=\"UTF-8\"><title>Extracted Images</title></head><body>");
            writer.WriteLine("<h1>Extracted Images</h1>");

            foreach (string imgPath in extractedImagePaths)
            {
                string relativePath = Path.GetFileName(imgPath);
                writer.WriteLine("<div style=\"margin-bottom:20px;\">");
                writer.WriteLine($"<p>{relativePath}</p>");
                writer.WriteLine($"<img src=\"{relativePath}\" style=\"max-width:600px; height:auto;\"/>");
                writer.WriteLine("</div>");
            }

            writer.WriteLine("</body></html>");
        }

        // Copy images to the same folder as the HTML file so that the <img> src works without subfolders.
        foreach (string imgPath in extractedImagePaths)
        {
            string destPath = Path.Combine(rootFolder, Path.GetFileName(imgPath));
            File.Copy(imgPath, destPath, true);
        }

        // The example finishes execution here. All files are written to the 'BatchImageExtraction' folder.
    }

    // Creates a deterministic PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.White);
                // Draw a simple rectangle to make the image non‑blank.
                using (Pen pen = new Pen(Aspose.Drawing.Color.Blue, 5))
                {
                    graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }

    // Creates a DOCX file that contains the specified image.
    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln($"Document generated for image extraction: {Path.GetFileName(docPath)}");
        builder.InsertImage(imagePath);
        doc.Save(docPath, SaveFormat.Docx);
    }
}
