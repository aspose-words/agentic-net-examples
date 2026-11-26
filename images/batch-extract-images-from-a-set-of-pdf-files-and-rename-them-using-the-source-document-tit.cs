using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputPdfs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputImages");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample PDF files with images and titles
        CreateSamplePdf(Path.Combine(inputFolder, "SampleDoc1.pdf"), "DocumentOne", Color.Blue);
        CreateSamplePdf(Path.Combine(inputFolder, "SampleDoc2.pdf"), "DocumentTwo", Color.Green);

        // Batch process PDFs
        string[] pdfFiles = Directory.GetFiles(inputFolder, "*.pdf");
        foreach (string pdfPath in pdfFiles)
        {
            // Load PDF as Aspose.Words document
            Document doc = new Document(pdfPath);
            string title = doc.BuiltInDocumentProperties.Title;
            if (string.IsNullOrEmpty(title))
                title = Path.GetFileNameWithoutExtension(pdfPath);

            // Find all image shapes
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    string imageFileName = $"{title}_{imageIndex}.png";
                    string imagePath = Path.Combine(outputFolder, imageFileName);
                    shape.ImageData.Save(imagePath);
                    if (!File.Exists(imagePath))
                        throw new Exception($"Failed to save extracted image to '{imagePath}'.");
                    imageIndex++;
                }
            }

            // Validate that at least one image was extracted
            if (imageIndex == 0)
                throw new Exception($"No images were extracted from document '{pdfPath}'.");
        }
    }

    private static void CreateSamplePdf(string pdfPath, string title, Color fillColor)
    {
        // Create a deterministic sample image using Aspose.Drawing
        string tempImagePath = Path.Combine(Path.GetDirectoryName(pdfPath) ?? "", "temp_image.png");
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(fillColor);
            }
            bitmap.Save(tempImagePath);
        }

        // Create a new Word document, set title, insert image, and save as PDF
        Document doc = new Document();
        doc.BuiltInDocumentProperties.Title = title;
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(tempImagePath);
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Clean up temporary image file
        if (File.Exists(tempImagePath))
            File.Delete(tempImagePath);
    }
}
