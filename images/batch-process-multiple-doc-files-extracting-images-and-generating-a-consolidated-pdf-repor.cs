using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Base working directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string imageDir = Path.Combine(baseDir, "ExtractedImages");
        string outputDir = Path.Combine(baseDir, "Output");

        // Ensure clean directories.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(imageDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic sample image (sample.png).
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 150);

        // -----------------------------------------------------------------
        // 2. Create sample DOCX files that contain the sample image.
        // -----------------------------------------------------------------
        int numberOfDocs = 3;
        for (int i = 1; i <= numberOfDocs; i++)
        {
            string docPath = Path.Combine(inputDir, $"Document{i}.docx");
            CreateSampleDocumentWithImage(docPath, sampleImagePath, $"Sample document {i}");
        }

        // -----------------------------------------------------------------
        // 3. Batch process: extract images from each DOCX.
        // -----------------------------------------------------------------
        var extractedImagePaths = new List<string>();

        foreach (string docFile in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(docFile);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

            int imageIndex = 0;
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    // Determine file extension based on image type.
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_img{imageIndex}{extension}";
                    string imagePath = Path.Combine(imageDir, imageFileName);

                    // Save the image to the file system.
                    shape.ImageData.Save(imagePath);
                    extractedImagePaths.Add(imagePath);
                    imageIndex++;
                }
            }
        }

        // Validate that at least one image was extracted.
        if (extractedImagePaths.Count == 0)
            throw new InvalidOperationException("No images were extracted from the source documents.");

        // -----------------------------------------------------------------
        // 4. Create a consolidated PDF report containing all extracted images.
        // -----------------------------------------------------------------
        Document reportDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(reportDoc);

        builder.Writeln("Consolidated Image Report");
        builder.Writeln("--------------------------");
        builder.Writeln();

        foreach (string imgPath in extractedImagePaths)
        {
            builder.Writeln($"Image from source: {Path.GetFileName(imgPath)}");
            builder.InsertImage(imgPath);
            builder.Writeln();
            builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the report as PDF with JPEG compression.
        string reportPdfPath = Path.Combine(outputDir, "ConsolidatedReport.pdf");
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ImageCompression = PdfImageCompression.Jpeg,
            JpegQuality = 80
        };
        reportDoc.Save(reportPdfPath, pdfOptions);

        // Validate that the PDF report was created.
        if (!File.Exists(reportPdfPath))
            throw new FileNotFoundException("Failed to create the consolidated PDF report.", reportPdfPath);
    }

    // Creates a simple white bitmap with a black rectangle.
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.White);
            using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black, 3))
            {
                graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
            }
            bitmap.Save(filePath);
        }
    }

    // Creates a DOCX file with a title and inserts the specified image.
    private static void CreateSampleDocumentWithImage(string docPath, string imagePath, string title)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(title);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }
}
