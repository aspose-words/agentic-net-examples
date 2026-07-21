using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class ExportHeaderFooterImages
{
    public static void Main()
    {
        // Prepare directories.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(baseDir);
        string headerImagesDir = Path.Combine(baseDir, "HeaderImages");
        string footerImagesDir = Path.Combine(baseDir, "FooterImages");
        Directory.CreateDirectory(headerImagesDir);
        Directory.CreateDirectory(footerImagesDir);

        // Create deterministic sample images for header and footer.
        string headerImagePath = Path.Combine(baseDir, "header.png");
        string footerImagePath = Path.Combine(baseDir, "footer.png");
        CreateSampleImage(headerImagePath, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage(footerImagePath, Aspose.Drawing.Color.LightGreen);

        // Build a sample ODT document with header and footer containing the images.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert header image.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.InsertImage(headerImagePath);

        // Insert footer image.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.InsertImage(footerImagePath);

        // Save the document as ODT.
        string odtPath = Path.Combine(baseDir, "SampleDocument.odt");
        doc.Save(odtPath, SaveFormat.Odt);

        // Load the document back (simulating a separate extraction step).
        Document loadedDoc = new Document(odtPath);

        // Extract images from headers.
        int headerCount = ExtractImagesFromHeaderFooters(loadedDoc, HeaderFooterType.HeaderPrimary, headerImagesDir);
        // Extract images from footers.
        int footerCount = ExtractImagesFromHeaderFooters(loadedDoc, HeaderFooterType.FooterPrimary, footerImagesDir);

        // Validate that images were extracted.
        if (headerCount == 0)
            throw new InvalidOperationException("No header images were extracted.");
        if (footerCount == 0)
            throw new InvalidOperationException("No footer images were extracted.");

        // Example completed successfully.
        Console.WriteLine($"Extracted {headerCount} header image(s) to: {headerImagesDir}");
        Console.WriteLine($"Extracted {footerCount} footer image(s) to: {footerImagesDir}");
    }

    // Creates a simple bitmap, fills it with a solid color, and saves it to the given path.
    private static void CreateSampleImage(string filePath, Aspose.Drawing.Color fillColor)
    {
        const int width = 100;
        const int height = 50;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(fillColor);
            }
            bitmap.Save(filePath);
        }
    }

    // Extracts images from all header/footer sections of the specified type and saves them to the target folder.
    private static int ExtractImagesFromHeaderFooters(Document doc, HeaderFooterType targetType, string outputFolder)
    {
        int savedCount = 0;
        foreach (Section section in doc.Sections)
        {
            HeaderFooter hf = section.HeadersFooters[targetType];
            if (hf == null)
                continue;

            NodeCollection shapes = hf.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            foreach (Shape shape in shapes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string fileName = Path.Combine(outputFolder, $"image_{imageIndex}{extension}");
                    shape.ImageData.Save(fileName);
                    imageIndex++;
                    savedCount++;
                }
            }
        }
        return savedCount;
    }
}
