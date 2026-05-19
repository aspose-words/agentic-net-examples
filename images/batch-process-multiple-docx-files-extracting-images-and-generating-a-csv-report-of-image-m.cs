using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Root folder for all temporary data.
        string rootFolder = Path.Combine(Directory.GetCurrentDirectory(), "BatchImageDemo");
        Directory.CreateDirectory(rootFolder);

        // Folder that will contain the generated DOCX files.
        string docsFolder = Path.Combine(rootFolder, "Docs");
        Directory.CreateDirectory(docsFolder);

        // Folder that will contain the extracted images.
        string imagesFolder = Path.Combine(rootFolder, "ExtractedImages");
        Directory.CreateDirectory(imagesFolder);

        // Folder for the CSV report.
        string reportFolder = Path.Combine(rootFolder, "Report");
        Directory.CreateDirectory(reportFolder);
        string csvPath = Path.Combine(reportFolder, "ImageMetadataReport.csv");

        // -----------------------------------------------------------------
        // 1. Create sample images (deterministic local files).
        // -----------------------------------------------------------------
        string sampleImage1 = Path.Combine(rootFolder, "Sample1.png");
        string sampleImage2 = Path.Combine(rootFolder, "Sample2.png");
        CreateSampleImage(sampleImage1, 200, 150, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage(sampleImage2, 300, 100, Aspose.Drawing.Color.LightCoral);

        // -----------------------------------------------------------------
        // 2. Create a few sample DOCX files that embed the images.
        // -----------------------------------------------------------------
        CreateSampleDocument(Path.Combine(docsFolder, "DocumentA.docx"), new[] { sampleImage1, sampleImage2 });
        CreateSampleDocument(Path.Combine(docsFolder, "DocumentB.docx"), new[] { sampleImage2 });

        // -----------------------------------------------------------------
        // 3. Batch process all DOCX files: extract images and collect metadata.
        // -----------------------------------------------------------------
        var csvLines = new List<string>();
        // Header row.
        csvLines.Add("Document,ImageFileName,ImageType,WidthPoints,HeightPoints,HorizontalResolutionDPI,VerticalResolutionDPI");

        int totalExtractedImages = 0;

        foreach (string docPath in Directory.GetFiles(docsFolder, "*.docx"))
        {
            Document doc = new Document(docPath);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                // Determine file extension based on the image type stored in the shape.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_Img{imageIndex}{extension}";
                string imageFullPath = Path.Combine(imagesFolder, imageFileName);

                // Save the image to the file system.
                shape.ImageData.Save(imageFullPath);

                // Gather metadata.
                ImageSize size = shape.ImageData.ImageSize;
                string line = string.Join(",",
                    Path.GetFileName(docPath),
                    imageFileName,
                    shape.ImageData.ImageType.ToString(),
                    size.WidthPoints.ToString("F2"),
                    size.HeightPoints.ToString("F2"),
                    size.HorizontalResolution.ToString("F2"),
                    size.VerticalResolution.ToString("F2"));

                csvLines.Add(line);
                imageIndex++;
                totalExtractedImages++;
            }
        }

        // Validate that at least one image was extracted.
        if (totalExtractedImages == 0)
            throw new InvalidOperationException("No images were extracted from the DOCX files.");

        // -----------------------------------------------------------------
        // 4. Write the CSV report.
        // -----------------------------------------------------------------
        File.WriteAllLines(csvPath, csvLines);

        // -----------------------------------------------------------------
        // 5. Clean up sample images used as inputs (optional).
        // -----------------------------------------------------------------
        File.Delete(sampleImage1);
        File.Delete(sampleImage2);
    }

    // Creates a deterministic PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color fillColor)
    {
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                graphics.Clear(fillColor);
            }
            bitmap.Save(filePath);
        }
    }

    // Creates a DOCX document and inserts the provided images sequentially.
    private static void CreateSampleDocument(string docPath, string[] imagePaths)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        foreach (string imgPath in imagePaths)
        {
            // Ensure the image file exists before insertion.
            if (!File.Exists(imgPath))
                throw new FileNotFoundException($"Image file not found: {imgPath}");

            // Insert the image inline.
            builder.InsertImage(imgPath);
            // Add a line break between images for readability.
            builder.Writeln();
        }

        doc.Save(docPath);
    }
}
