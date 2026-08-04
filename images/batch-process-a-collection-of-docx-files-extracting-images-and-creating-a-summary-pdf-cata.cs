using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Loading;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Base directories for the example.
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string extractedDir = Path.Combine(baseDir, "ExtractedImages");
        string outputDir = Path.Combine(baseDir, "OutputPdf");

        // Ensure directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(extractedDir);
        Directory.CreateDirectory(outputDir);

        // Create a deterministic sample image (sample.png).
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSampleImage(sampleImagePath, 100, 100);

        // Create a few sample DOCX files that contain the sample image.
        CreateSampleDocuments(inputDir, sampleImagePath, 3);

        // List to hold paths of all extracted images for the summary PDF.
        List<string> allExtractedImages = new List<string>();

        // Batch process each DOCX file in the input directory.
        foreach (string docPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(docPath);
            var shapes = doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>();

            int imageIndex = 0;
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string extractedImagePath = Path.Combine(
                        extractedDir,
                        $"{Path.GetFileNameWithoutExtension(docPath)}_Image_{imageIndex}{extension}");

                    shape.ImageData.Save(extractedImagePath);
                    allExtractedImages.Add(extractedImagePath);
                    imageIndex++;
                }
            }

            // Validation: at least one image should have been extracted from each document.
            if (imageIndex == 0)
                throw new InvalidOperationException($"No images were extracted from '{docPath}'.");
        }

        // Create a summary PDF catalog that contains all extracted images.
        Document summaryDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(summaryDoc);

        foreach (string imagePath in allExtractedImages)
        {
            builder.Writeln(Path.GetFileName(imagePath));
            builder.InsertImage(imagePath);
            builder.InsertParagraph();
        }

        // Save the summary as PDF with JPEG compression.
        string pdfPath = Path.Combine(outputDir, "SummaryCatalog.pdf");
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ImageCompression = PdfImageCompression.Jpeg,
            JpegQuality = 80
        };
        summaryDoc.Save(pdfPath, pdfOptions);

        // Validation: ensure the PDF was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("The summary PDF catalog was not created.", pdfPath);

        Console.WriteLine("Batch image extraction and PDF catalog creation completed successfully.");
    }

    // Creates a deterministic PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill the background with white.
                graphics.Clear(Aspose.Drawing.Color.White);
                // Draw a simple black rectangle for visual distinction.
                using (Pen pen = new Pen(Aspose.Drawing.Color.Black, 2))
                {
                    graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            bitmap.Save(filePath);
        }
    }

    // Generates a specified number of DOCX files, each containing the sample image.
    private static void CreateSampleDocuments(string folderPath, string imagePath, int count)
    {
        for (int i = 1; i <= count; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln($"Sample Document {i}");
            // Insert the sample image twice to have multiple images per document.
            builder.InsertImage(imagePath);
            builder.InsertParagraph();
            builder.InsertImage(imagePath);
            builder.InsertParagraph();

            string docFilePath = Path.Combine(folderPath, $"Doc{i}.docx");
            doc.Save(docFilePath);
        }
    }
}
