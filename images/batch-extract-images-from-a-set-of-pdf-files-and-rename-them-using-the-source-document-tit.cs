using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Directories for input PDFs and extracted images.
        string inputDir = Path.Combine(Directory.GetCurrentDirectory(), "InputPdfs");
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "ExtractedImages");

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample PDF files with images.
        CreateSamplePdfWithImage(Path.Combine(inputDir, "Sample1.pdf"), "SampleDoc1");
        CreateSamplePdfWithImage(Path.Combine(inputDir, "Sample2.pdf"), "SampleDoc2");

        // Process each PDF in the input directory.
        foreach (string pdfPath in Directory.GetFiles(inputDir, "*.pdf"))
        {
            // Load the PDF document.
            Document pdfDoc = new Document(pdfPath, new LoadOptions());

            // Retrieve the document title; fall back to file name without extension.
            string title = pdfDoc.BuiltInDocumentProperties.Title;
            if (string.IsNullOrWhiteSpace(title))
                title = Path.GetFileNameWithoutExtension(pdfPath);

            int imageIndex = 0;

            // Iterate over all Shape nodes that contain images.
            var shapes = pdfDoc.GetChildNodes(NodeType.Shape, true)
                               .OfType<Shape>()
                               .Where(s => s.HasImage);

            foreach (Shape shape in shapes)
            {
                // Determine the appropriate file extension for the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{title}_Image{imageIndex}{extension}";
                string imagePath = Path.Combine(outputDir, imageFileName);

                // Save the image to the output folder.
                shape.ImageData.Save(imagePath);
                imageIndex++;
            }

            // Validation: ensure at least one image was extracted.
            if (imageIndex == 0)
                throw new InvalidOperationException($"No images were extracted from '{pdfPath}'.");
        }
    }

    // Creates a PDF file containing a single image and sets the document title.
    private static void CreateSamplePdfWithImage(string pdfPath, string title)
    {
        // Create a deterministic sample image.
        string tempImagePath = Path.ChangeExtension(pdfPath, ".png");
        CreateSampleImage(tempImagePath);

        // Build a simple document with a title and the image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the document title property.
        doc.BuiltInDocumentProperties.Title = title;

        // Insert the title as a heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln(title);

        // Insert the sample image.
        builder.InsertImage(tempImagePath);

        // Save as PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Clean up the temporary image file.
        if (File.Exists(tempImagePath))
            File.Delete(tempImagePath);
    }

    // Generates a simple 100x100 PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string filePath)
    {
        const int width = 100;
        const int height = 100;

        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            // Fill the background with a solid color.
            graphics.Clear(Color.LightBlue);
            // Optionally, draw a simple rectangle.
            graphics.DrawRectangle(Pens.Black, 10, 10, width - 20, height - 20);
            // Save the bitmap to the specified path.
            bitmap.Save(filePath);
        }
    }
}
