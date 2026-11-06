using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Define folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string outputImageFolder = Path.Combine(baseDir, "ExtractedImages");
        string summaryCsvPath = Path.Combine(baseDir, "Summary.csv");

        // Ensure clean environment
        if (Directory.Exists(inputFolder)) Directory.Delete(inputFolder, true);
        if (Directory.Exists(outputImageFolder)) Directory.Delete(outputImageFolder, true);
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputImageFolder);

        // Create a deterministic sample image to be used in documents
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // Create sample DOCX files with images
        CreateSampleDocument(Path.Combine(inputFolder, "Doc1.docx"), sampleImagePath, 1);
        CreateSampleDocument(Path.Combine(inputFolder, "Doc2.docx"), sampleImagePath, 2);

        // Prepare CSV rows
        var csvRows = new List<string>();
        csvRows.Add("DocumentName,ImageFileName,Width,Height,SizeBytes");

        int totalExtractedImages = 0;

        // Process each DOCX file
        foreach (string docPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(docPath);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapes)
            {
                if (!shape.HasImage) continue;

                imageIndex++;
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_Image{imageIndex}.png";
                string imageFullPath = Path.Combine(outputImageFolder, imageFileName);

                // Save the image to disk
                shape.ImageData.Save(imageFullPath);

                // Load saved image to obtain dimensions and size
                using (Aspose.Drawing.Image img = Aspose.Drawing.Image.FromFile(imageFullPath))
                {
                    long fileSize = new FileInfo(imageFullPath).Length;
                    csvRows.Add($"{Path.GetFileName(docPath)},{imageFileName},{img.Width},{img.Height},{fileSize}");
                }

                totalExtractedImages++;
            }
        }

        // Validate that at least one image was extracted
        if (totalExtractedImages == 0)
            throw new InvalidOperationException("No images were extracted from the DOCX files.");

        // Write CSV summary
        File.WriteAllLines(summaryCsvPath, csvRows);
    }

    // Creates a simple white PNG image with given dimensions
    private static void CreateSampleImage(string path, int width, int height)
    {
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        bitmap.Save(path, Aspose.Drawing.Imaging.ImageFormat.Png);
        graphics.Dispose();
        bitmap.Dispose();
    }

    // Creates a DOCX file containing a specified number of the sample image
    private static void CreateSampleDocument(string docPath, string imagePath, int imageCount)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 0; i < imageCount; i++)
        {
            // Insert a paragraph break between images for clarity
            if (i > 0) builder.Writeln();

            // Insert the image
            builder.InsertImage(imagePath);
        }

        doc.Save(docPath);
    }
}
