using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Base working directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string docsDir = Path.Combine(baseDir, "Docs");
        string extractedImagesDir = Path.Combine(baseDir, "ExtractedImages");
        string indexFilePath = Path.Combine(baseDir, "ImageIndex.csv");
        string sampleImagePath = Path.Combine(baseDir, "sample.png");

        // Ensure required folders exist.
        Directory.CreateDirectory(docsDir);
        Directory.CreateDirectory(extractedImagesDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic sample image (required by the rules).
        // -----------------------------------------------------------------
        CreateSampleImage(sampleImagePath);

        // -----------------------------------------------------------------
        // 2. Generate a few sample DOCX files that contain the image.
        // -----------------------------------------------------------------
        const int sampleDocCount = 3;
        for (int i = 1; i <= sampleDocCount; i++)
        {
            string docPath = Path.Combine(docsDir, $"Doc{i}.docx");
            CreateDocumentWithImage(docPath, sampleImagePath);
        }

        // -----------------------------------------------------------------
        // 3. Batch extract images from all DOC/DOCX files and build index.
        // -----------------------------------------------------------------
        var indexLines = new List<string>();
        indexLines.Add("DocumentPath,ImagePath"); // CSV header

        string[] docFiles = Directory.GetFiles(docsDir, "*.*", SearchOption.TopDirectoryOnly);
        foreach (string docFile in docFiles)
        {
            // Load each document (lifecycle rule).
            Document doc = new Document(docFile);

            // Get all shape nodes.
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    // Determine proper file extension for the image type.
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_img{imageIndex}{extension}";
                    string imageFullPath = Path.Combine(extractedImagesDir, imageFileName);

                    // Save the image (extraction rule).
                    shape.ImageData.Save(imageFullPath);

                    // Record the mapping.
                    indexLines.Add($"{docFile},{imageFullPath}");
                    imageIndex++;
                }
            }
        }

        // Validation: ensure at least one image was extracted.
        if (indexLines.Count <= 1)
            throw new InvalidOperationException("No images were extracted from the documents.");

        // Write the CSV index (Excel can open CSV files).
        File.WriteAllLines(indexFilePath, indexLines);
    }

    // Creates a simple 100x100 PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string filePath)
    {
        const int width = 100;
        const int height = 100;

        Bitmap bitmap = new Bitmap(width, height);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.LightBlue);
        // Optional: draw a simple rectangle.
        graphics.DrawRectangle(new Pen(Color.DarkBlue, 2), 10, 10, 80, 80);
        bitmap.Save(filePath);
        graphics.Dispose();
        bitmap.Dispose();
    }

    // Creates a blank document and inserts the provided image.
    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        // Save the document (lifecycle rule).
        doc.Save(docPath);
    }
}
