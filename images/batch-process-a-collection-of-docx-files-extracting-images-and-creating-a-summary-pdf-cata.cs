using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing; // Provides Bitmap, Graphics, Color

public class Program
{
    // Entry point
    public static void Main()
    {
        // Define folders
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        string inputDocsDir = Path.Combine(artifactsDir, "InputDocs");
        string extractedImagesDir = Path.Combine(artifactsDir, "ExtractedImages");
        string summaryPdfPath = Path.Combine(artifactsDir, "SummaryCatalog.pdf");

        // Ensure clean environment
        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(inputDocsDir);
        Directory.CreateDirectory(extractedImagesDir);

        // 1. Create a deterministic sample image (sample.png)
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200, Aspose.Drawing.Color.LightBlue, Aspose.Drawing.Color.DarkBlue);

        // 2. Generate a few sample DOCX files that contain the sample image
        int docCount = 3;
        List<string> docPaths = new List<string>();
        for (int i = 1; i <= docCount; i++)
        {
            string docPath = Path.Combine(inputDocsDir, $"Document{i}.docx");
            CreateSampleDocumentWithImage(docPath, sampleImagePath, $"Sample document {i}");
            docPaths.Add(docPath);
        }

        // 3. Batch process each DOCX: extract images and collect extracted image paths per document
        var docToImagesMap = new Dictionary<string, List<string>>();
        foreach (string docPath in docPaths)
        {
            var extractedForDoc = ExtractImagesFromDocument(docPath, extractedImagesDir);
            if (extractedForDoc.Count == 0)
                throw new InvalidOperationException($"No images were extracted from '{docPath}'.");
            docToImagesMap[docPath] = extractedForDoc;
        }

        // 4. Create a summary PDF catalog that lists each document and shows its extracted images
        CreateSummaryPdfCatalog(docToImagesMap, summaryPdfPath);

        // 5. Validate final outputs
        if (!File.Exists(summaryPdfPath))
            throw new InvalidOperationException("Summary PDF catalog was not created.");

        Console.WriteLine("Batch processing completed successfully.");
        Console.WriteLine($"Extracted images are located in: {extractedImagesDir}");
        Console.WriteLine($"Summary PDF catalog: {summaryPdfPath}");
    }

    // Creates a simple bitmap with a colored rectangle and saves it to a file
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color background, Aspose.Drawing.Color rectangleColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(background);
            // Draw a filled rectangle in the centre
            int rectSize = Math.Min(width, height) / 2;
            int rectX = (width - rectSize) / 2;
            int rectY = (height - rectSize) / 2;
            graphics.FillRectangle(new SolidBrush(rectangleColor), rectX, rectY, rectSize, rectSize);
            bitmap.Save(filePath);
        }
    }

    // Creates a DOCX file containing a single image and a paragraph of text
    private static void CreateSampleDocumentWithImage(string docPath, string imagePath, string title)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(title);
        builder.InsertParagraph();
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Extracts all images from a DOCX and saves them to the target folder.
    // Returns a list of full file paths of the saved images.
    private static List<string> ExtractImagesFromDocument(string docPath, string targetFolder)
    {
        Document doc = new Document(docPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        List<string> savedImages = new List<string>();
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
            string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_image{imageIndex}{extension}";
            string fullPath = Path.Combine(targetFolder, imageFileName);
            shape.ImageData.Save(fullPath);
            savedImages.Add(fullPath);
            imageIndex++;
        }

        return savedImages;
    }

    // Builds a PDF catalog that lists each source document and embeds its extracted images.
    private static void CreateSummaryPdfCatalog(Dictionary<string, List<string>> docImagesMap, string outputPdfPath)
    {
        Document summaryDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(summaryDoc);

        foreach (var entry in docImagesMap)
        {
            string docName = Path.GetFileName(entry.Key);
            builder.Writeln($"Document: {docName}");
            builder.Writeln(); // empty line

            foreach (string imagePath in entry.Value)
            {
                // Insert each extracted image; scale it down to keep the catalog readable
                builder.InsertImage(imagePath, 150, 150);
                builder.Writeln(); // space between images
            }

            builder.Writeln(); // extra space before next document section
            builder.InsertBreak(BreakType.PageBreak);
        }

        // Save as PDF
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        summaryDoc.Save(outputPdfPath, pdfOptions);
    }
}
