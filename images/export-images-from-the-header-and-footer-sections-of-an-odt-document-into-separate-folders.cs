using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ExportHeaderFooterImages
{
    public static void Main()
    {
        // Prepare deterministic file and folder names
        const string headerImagePath = "header.png";
        const string footerImagePath = "footer.png";
        const string documentPath = "sample.odt";
        const string headerImagesFolder = "HeaderImages";
        const string footerImagesFolder = "FooterImages";

        // Create sample images using Aspose.Drawing
        CreateSampleImage(headerImagePath, 200, 50, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage(footerImagePath, 200, 50, Aspose.Drawing.Color.LightGreen);

        // Build a document with header and footer containing the images
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert image into the primary header
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.InsertImage(headerImagePath);

        // Insert image into the primary footer
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.InsertImage(footerImagePath);

        // Add some body content
        builder.MoveToDocumentEnd();
        builder.Writeln("This is the body of the document.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page content.");

        // Save the document as ODT
        doc.Save(documentPath, SaveFormat.Odt);

        // Load the saved document (demonstrates load rule)
        Document loadedDoc = new Document(documentPath);

        // Ensure output folders exist
        Directory.CreateDirectory(headerImagesFolder);
        Directory.CreateDirectory(footerImagesFolder);

        int headerImageCount = 0;
        int footerImageCount = 0;

        // Iterate through each section's headers and footers
        foreach (Section section in loadedDoc.Sections)
        {
            foreach (HeaderFooter hf in section.HeadersFooters)
            {
                // Determine target folder based on header/footer type
                bool isHeader = hf.HeaderFooterType.ToString().StartsWith("Header", StringComparison.Ordinal);
                string targetFolder = isHeader ? headerImagesFolder : footerImagesFolder;

                // Collect all Shape nodes that may contain images
                NodeCollection shapeNodes = hf.GetChildNodes(NodeType.Shape, true);
                foreach (Shape shape in shapeNodes.OfType<Shape>())
                {
                    if (shape.HasImage)
                    {
                        string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                        string fileName = Path.Combine(targetFolder,
                            $"extracted_{Guid.NewGuid()}{extension}");
                        shape.ImageData.Save(fileName);

                        if (isHeader)
                            headerImageCount++;
                        else
                            footerImageCount++;
                    }
                }
            }
        }

        // Validation: ensure at least one image was extracted from each part
        if (headerImageCount == 0)
            throw new InvalidOperationException("No images were extracted from headers.");
        if (footerImageCount == 0)
            throw new InvalidOperationException("No images were extracted from footers.");

        // Cleanup sample images (optional)
        File.Delete(headerImagePath);
        File.Delete(footerImagePath);
    }

    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backgroundColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(backgroundColor);
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}
