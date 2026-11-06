using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare deterministic folders
        const string headerFolder = "HeaderImages";
        const string footerFolder = "FooterImages";
        Directory.CreateDirectory(headerFolder);
        Directory.CreateDirectory(footerFolder);

        // Create sample images for header and footer
        const string headerImagePath = "header.png";
        const string footerImagePath = "footer.png";

        CreateSampleImage(headerImagePath, 200, 100, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage(footerImagePath, 200, 100, Aspose.Drawing.Color.LightGreen);

        // Create a new ODT document and insert the images into header and footer
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert image into primary header
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.InsertImage(headerImagePath);

        // Insert image into primary footer
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.InsertImage(footerImagePath);

        // Save the document as ODT
        const string docPath = "sample.odt";
        doc.Save(docPath, SaveFormat.Odt);

        // Reload the document (simulating a separate operation)
        Document loadedDoc = new Document(docPath);

        int headerImageCount = 0;
        int footerImageCount = 0;

        // Iterate through all sections and their header/footer collections
        foreach (Section section in loadedDoc.Sections)
        {
            foreach (HeaderFooter headerFooter in section.HeadersFooters)
            {
                bool isHeader = headerFooter.HeaderFooterType.ToString().Contains("Header");
                NodeCollection shapeNodes = headerFooter.GetChildNodes(NodeType.Shape, true);

                foreach (Shape shape in shapeNodes)
                {
                    if (!shape.HasImage) continue;

                    string outputPath;
                    if (isHeader)
                    {
                        outputPath = Path.Combine(headerFolder, $"header-{++headerImageCount}.png");
                    }
                    else
                    {
                        outputPath = Path.Combine(footerFolder, $"footer-{++footerImageCount}.png");
                    }

                    shape.ImageData.Save(outputPath);
                }
            }
        }

        // Validation
        if (headerImageCount == 0)
            throw new InvalidOperationException("No header images were extracted.");
        if (footerImageCount == 0)
            throw new InvalidOperationException("No footer images were extracted.");

        Console.WriteLine($"Extraction complete. Header images: {headerImageCount}, Footer images: {footerImageCount}");
    }

    // Helper method to create a deterministic sample image
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(backColor);
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}
