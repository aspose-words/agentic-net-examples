using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Loading;
using Aspose.Drawing; // Provides Bitmap, Graphics, Color, Pen

public class Program
{
    public static void Main()
    {
        // Base directories
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string extractedImagesDir = Path.Combine(baseDir, "ExtractedImages");
        string summaryPdfPath = Path.Combine(baseDir, "ImageCatalog.pdf");
        string sampleImagePath = Path.Combine(baseDir, "sample.png");

        // Ensure clean folders
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(extractedImagesDir);

        // -------------------------------------------------
        // 1. Create a deterministic sample image (sample.png)
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imgWidth, imgHeight))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                // Fill background with white
                g.Clear(Aspose.Drawing.Color.White);
                // Draw a simple black rectangle
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black, 3))
                {
                    g.DrawRectangle(pen, 10, 10, imgWidth - 20, imgHeight - 20);
                }
            }
            // Save the image to a file that will be used for insertion
            bitmap.Save(sampleImagePath);
        }

        // -------------------------------------------------
        // 2. Create sample DOCX files that contain the image
        // -------------------------------------------------
        const int docCount = 3;
        for (int i = 1; i <= docCount; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Document {i} - contains a sample image.");
            // Insert the previously created sample image
            builder.InsertImage(sampleImagePath);
            builder.Writeln("End of document.");
            string docPath = Path.Combine(inputDir, $"Doc{i}.docx");
            doc.Save(docPath);
        }

        // -------------------------------------------------
        // 3. Batch process each DOCX: extract images
        // -------------------------------------------------
        var extractedImagePaths = new List<(string SourceDoc, string ImagePath)>();

        foreach (string docFile in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(docFile);
            // Get all shape nodes (including images)
            var shapes = doc.GetChildNodes(NodeType.Shape, true)
                            .Cast<Shape>()
                            .Where(s => s.HasImage);

            int imageIndex = 0;
            foreach (Shape shape in shapes)
            {
                // Determine proper file extension based on image type
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_Image{imageIndex}{extension}";
                string imageFullPath = Path.Combine(extractedImagesDir, imageFileName);

                // Save the image to the file system
                shape.ImageData.Save(imageFullPath);
                extractedImagePaths.Add((Path.GetFileName(docFile), imageFullPath));
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted
        if (!extractedImagePaths.Any())
            throw new InvalidOperationException("No images were extracted from the input documents.");

        // -------------------------------------------------
        // 4. Create a summary PDF catalog containing all extracted images
        // -------------------------------------------------
        Document summaryDoc = new Document();
        DocumentBuilder summaryBuilder = new DocumentBuilder(summaryDoc);

        // Group images by source document for clearer catalog layout
        var imagesByDoc = extractedImagePaths
                         .GroupBy(item => item.SourceDoc)
                         .OrderBy(g => g.Key);

        foreach (var group in imagesByDoc)
        {
            // Add a heading for each source document
            summaryBuilder.Writeln($"Images extracted from {group.Key}:");
            summaryBuilder.Font.Size = 12;
            summaryBuilder.Font.Bold = true;
            summaryBuilder.Writeln();

            int imgNum = 0;
            foreach (var (sourceDoc, imagePath) in group)
            {
                // Insert the image into the catalog
                summaryBuilder.InsertImage(imagePath);
                summaryBuilder.Writeln($"Image {imgNum + 1} from {sourceDoc}");
                summaryBuilder.Writeln(); // Add spacing
                imgNum++;
            }

            // Add a page break after each document's images (except after the last)
            if (group != imagesByDoc.Last())
                summaryBuilder.InsertBreak(BreakType.PageBreak);
        }

        // Save the catalog as PDF
        summaryDoc.Save(summaryPdfPath, SaveFormat.Pdf);

        // Validate that the PDF catalog was created
        if (!File.Exists(summaryPdfPath))
            throw new InvalidOperationException("PDF catalog was not created.");
    }
}
