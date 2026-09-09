using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Tables;
using Aspose.Drawing;

public class BatchImageExtractor
{
    // Paths used by the example
    private const string InputFolder = "InputDocs";
    private const string ImageOutputFolder = "ExtractedImages";
    private const string IndexFilePath = "ImageIndex.xlsx";

    public static void Main()
    {
        // Ensure clean environment
        PrepareFolders();

        // Step 1: Create sample DOCX files with images
        CreateSampleDocuments();

        // Step 2: Extract images from each document and collect index data
        var indexEntries = new List<(string DocumentPath, string ImagePath)>();
        foreach (string docPath in Directory.GetFiles(InputFolder, "*.docx"))
        {
            var extractedImages = ExtractImagesFromDocument(docPath);
            foreach (string imgPath in extractedImages)
            {
                indexEntries.Add((docPath, imgPath));
            }
        }

        // Validate that we extracted at least one image
        if (indexEntries.Count == 0)
            throw new InvalidOperationException("No images were extracted from the documents.");

        // Step 3: Create an Excel (XLSX) index using Aspose.Words
        CreateExcelIndex(indexEntries);

        // Validate that the index file was created
        if (!File.Exists(IndexFilePath))
            throw new InvalidOperationException($"Failed to create the index file at '{IndexFilePath}'.");

        Console.WriteLine("Image extraction and index generation completed successfully.");
    }

    private static void PrepareFolders()
    {
        // Delete previous runs (if any) and recreate folders
        if (Directory.Exists(InputFolder))
            Directory.Delete(InputFolder, true);
        if (Directory.Exists(ImageOutputFolder))
            Directory.Delete(ImageOutputFolder, true);

        Directory.CreateDirectory(InputFolder);
        Directory.CreateDirectory(ImageOutputFolder);
    }

    private static void CreateSampleDocuments()
    {
        // Create three sample images
        string[] sampleImageFiles = new string[3];
        for (int i = 0; i < 3; i++)
        {
            string imgPath = Path.Combine(InputFolder, $"sample{i + 1}.png");
            CreateSampleImage(imgPath, 100 + i * 50, 100 + i * 50, GetColor(i));
            sampleImageFiles[i] = imgPath;
        }

        // Create three DOCX files, each containing one of the sample images
        for (int i = 0; i < 3; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Document {i + 1} with an image:");
            builder.InsertImage(sampleImageFiles[i]);
            string docPath = Path.Combine(InputFolder, $"Document{i + 1}.docx");
            doc.Save(docPath);
        }
    }

    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color color)
    {
        // Create a bitmap, fill it with a solid color, and save as PNG
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(color);
            bitmap.Save(filePath);
        }
    }

    private static Aspose.Drawing.Color GetColor(int index)
    {
        // Return deterministic colors for the sample images
        return index switch
        {
            0 => Aspose.Drawing.Color.Red,
            1 => Aspose.Drawing.Color.Green,
            2 => Aspose.Drawing.Color.Blue,
            _ => Aspose.Drawing.Color.Black,
        };
    }

    private static List<string> ExtractImagesFromDocument(string docPath)
    {
        var extractedPaths = new List<string>();
        Document doc = new Document(docPath);

        // Get all shape nodes (including images)
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Determine file extension based on image type
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
            string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_Image{imageIndex}{extension}";
            string imageFullPath = Path.Combine(ImageOutputFolder, imageFileName);

            // Save the image to disk
            shape.ImageData.Save(imageFullPath);
            extractedPaths.Add(imageFullPath);
            imageIndex++;
        }

        return extractedPaths;
    }

    private static void CreateExcelIndex(List<(string DocumentPath, string ImagePath)> entries)
    {
        // Create a new Word document that will be saved as XLSX
        Document indexDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(indexDoc);

        // Insert a table with two columns: Document and Image File
        Table table = builder.StartTable();

        // Header row
        builder.InsertCell();
        builder.Write("Document");
        builder.InsertCell();
        builder.Write("Image File");
        builder.EndRow();

        // Data rows
        foreach (var entry in entries)
        {
            builder.InsertCell();
            builder.Write(Path.GetFileName(entry.DocumentPath));
            builder.InsertCell();
            builder.Write(Path.GetFileName(entry.ImagePath));
            builder.EndRow();
        }

        builder.EndTable();

        // Save the table as an Excel workbook
        indexDoc.Save(IndexFilePath, SaveFormat.Xlsx);
    }
}
