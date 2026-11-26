using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string inputFolder = "InputDocs";
        string imageOutputFolder = "ExtractedImages";
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(imageOutputFolder);

        // Create a deterministic sample image (input.png)
        string sampleImagePath = "input.png";
        CreateSampleImage(sampleImagePath, 200, 200);

        // Create sample DOCX files that contain the sample image
        CreateSampleDocument(Path.Combine(inputFolder, "Document1.docx"), sampleImagePath);
        CreateSampleDocument(Path.Combine(inputFolder, "Document2.docx"), sampleImagePath);

        // List to hold CSV rows
        var csvRows = new List<string>();
        csvRows.Add("Document,ImageFile,Width,Height,SizeBytes");

        // Process each DOCX file in the input folder
        foreach (string docPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(docPath);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapes)
            {
                if (!shape.HasImage) continue;

                // Build deterministic image file name
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_image_{imageIndex}.png";
                string imageFullPath = Path.Combine(imageOutputFolder, imageFileName);

                // Save the image from the shape
                shape.ImageData.Save(imageFullPath);

                // Load saved image to obtain dimensions
                using (Image img = Image.FromFile(imageFullPath))
                {
                    int width = img.Width;
                    int height = img.Height;
                    long sizeBytes = new FileInfo(imageFullPath).Length;

                    // Add row to CSV data
                    csvRows.Add($"{Path.GetFileName(docPath)},{imageFileName},{width},{height},{sizeBytes}");
                }

                // Validate that the image file was created
                if (!File.Exists(imageFullPath))
                    throw new Exception($"Failed to create image file: {imageFullPath}");

                imageIndex++;
            }
        }

        // Validate that at least one image was extracted
        if (csvRows.Count <= 1)
            throw new Exception("No images were extracted from the documents.");

        // Write summary CSV
        string csvPath = "summary.csv";
        File.WriteAllLines(csvPath, csvRows, Encoding.UTF8);

        // Validate CSV creation
        if (!File.Exists(csvPath))
            throw new Exception($"Failed to create summary CSV at {csvPath}");
    }

    // Creates a simple white PNG image with given dimensions
    private static void CreateSampleImage(string path, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                // Optional: draw a simple rectangle for visual distinction
                using (Pen pen = new Pen(Aspose.Drawing.Color.Black))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            bitmap.Save(path, ImageFormat.Png);
        }

        // Validate that the image file exists
        if (!File.Exists(path))
            throw new Exception($"Failed to create sample image at {path}");
    }

    // Creates a DOCX file that inserts the specified image
    private static void CreateSampleDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln($"This document contains an image inserted from {Path.GetFileName(imagePath)}:");
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        // Validate that the document file exists
        if (!File.Exists(docPath))
            throw new Exception($"Failed to create sample document at {docPath}");
    }
}
