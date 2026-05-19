using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Loading;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Base working directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string imageDir = Path.Combine(baseDir, "Images");
        string extractedDir = Path.Combine(baseDir, "ExtractedImages");
        string outputDir = Path.Combine(baseDir, "Output");

        // Ensure all directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(imageDir);
        Directory.CreateDirectory(extractedDir);
        Directory.CreateDirectory(outputDir);

        // -------------------------------------------------
        // 1. Create a deterministic sample image (sample.png).
        // -------------------------------------------------
        string sampleImagePath = Path.Combine(imageDir, "sample.png");
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Simple visual cue: a filled rectangle.
                g.FillRectangle(new SolidBrush(Color.Blue), 50, 50, 100, 100);
            }
            bitmap.Save(sampleImagePath);
        }

        // -------------------------------------------------
        // 2. Create several sample DOCX files that contain images.
        // -------------------------------------------------
        int docCount = 3;
        for (int i = 1; i <= docCount; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln($"Document {i}");
            // Insert the sample image twice in each document.
            builder.InsertImage(sampleImagePath);
            builder.InsertParagraph();
            builder.InsertImage(sampleImagePath);

            string docPath = Path.Combine(inputDir, $"Doc{i}.docx");
            doc.Save(docPath);
        }

        // -------------------------------------------------
        // 3. Batch process all DOCX files: extract images.
        // -------------------------------------------------
        var extractedImagePaths = new System.Collections.Generic.List<string>();
        string[] docFiles = Directory.GetFiles(inputDir, "*.docx");
        if (docFiles.Length == 0)
            throw new InvalidOperationException("No DOCX files found for processing.");

        int docIndex = 0;
        foreach (string docFile in docFiles)
        {
            docIndex++;
            Document doc = new Document(docFile);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"Doc{docIndex}_Img{++imageIndex}{extension}";
                string imagePath = Path.Combine(extractedDir, imageFileName);
                shape.ImageData.Save(imagePath);
                extractedImagePaths.Add(imagePath);
            }
        }

        // Validate that at least one image was extracted.
        if (extractedImagePaths.Count == 0)
            throw new InvalidOperationException("No images were extracted from the source documents.");

        // -------------------------------------------------
        // 4. Generate a consolidated PDF report containing all extracted images.
        // -------------------------------------------------
        Document pdfReport = new Document();
        DocumentBuilder pdfBuilder = new DocumentBuilder(pdfReport);

        foreach (string imgPath in extractedImagePaths)
        {
            // Insert a page break before each new image except the first.
            if (pdfBuilder.CurrentParagraph != null && pdfBuilder.CurrentParagraph.GetText().Trim().Length > 0)
                pdfBuilder.InsertBreak(BreakType.PageBreak);

            pdfBuilder.InsertImage(imgPath);
        }

        string pdfPath = Path.Combine(outputDir, "ConsolidatedReport.pdf");
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfReport.Save(pdfPath, pdfOptions);

        // -------------------------------------------------
        // 5. Final validation.
        // -------------------------------------------------
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the consolidated PDF report.");

        // The program finishes without requiring user interaction.
    }
}
