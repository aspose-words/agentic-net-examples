using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class ExportHeaderFooterImages
{
    public static void Main()
    {
        // Prepare output folders.
        string headerImagesFolder = Path.Combine(Directory.GetCurrentDirectory(), "HeaderImages");
        string footerImagesFolder = Path.Combine(Directory.GetCurrentDirectory(), "FooterImages");
        Directory.CreateDirectory(headerImagesFolder);
        Directory.CreateDirectory(footerImagesFolder);

        // Create deterministic sample images for header and footer.
        string headerImagePath = Path.Combine(Directory.GetCurrentDirectory(), "header.png");
        string footerImagePath = Path.Combine(Directory.GetCurrentDirectory(), "footer.png");
        CreateSampleImage(headerImagePath, 100, 100, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage(footerImagePath, 100, 100, Aspose.Drawing.Color.LightGreen);

        // Build a sample ODT document with images in header and footer.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert image into the primary header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.InsertImage(headerImagePath);

        // Insert image into the primary footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.InsertImage(footerImagePath);

        // Add body content.
        builder.MoveToDocumentEnd();
        builder.Writeln("Sample document body text.");

        // Save the document as ODT.
        string odtPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleDocument.odt");
        doc.Save(odtPath, SaveFormat.Odt);

        // Load the saved ODT document for extraction.
        Document loadedDoc = new Document(odtPath);

        int headerImageCount = 0;
        int footerImageCount = 0;

        // Iterate through all sections and their headers/footers.
        foreach (Section section in loadedDoc.Sections)
        {
            foreach (HeaderFooter headerFooter in section.HeadersFooters)
            {
                // Collect all Shape nodes inside the header/footer.
                NodeCollection shapes = headerFooter.GetChildNodes(NodeType.Shape, true);
                foreach (Shape shape in shapes)
                {
                    if (!shape.HasImage) continue;

                    // Determine target folder based on header/footer type.
                    bool isHeader = IsHeader(headerFooter.HeaderFooterType);
                    string targetFolder = isHeader ? headerImagesFolder : footerImagesFolder;
                    int index = isHeader ? ++headerImageCount : ++footerImageCount;

                    // Build deterministic file name with proper extension.
                    string fileName = $"Image_{index}{FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType)}";
                    string fullPath = Path.Combine(targetFolder, fileName);

                    // Save the image.
                    shape.ImageData.Save(fullPath);
                }
            }
        }

        // Validation: ensure at least one image was extracted for each part.
        if (headerImageCount == 0)
            throw new InvalidOperationException("No header images were extracted.");
        if (footerImageCount == 0)
            throw new InvalidOperationException("No footer images were extracted.");

        // Cleanup temporary files (optional).
        File.Delete(headerImagePath);
        File.Delete(footerImagePath);
    }

    // Helper to decide if a HeaderFooterType represents a header.
    private static bool IsHeader(HeaderFooterType type)
    {
        return type == HeaderFooterType.HeaderPrimary ||
               type == HeaderFooterType.HeaderFirst ||
               type == HeaderFooterType.HeaderEven;
    }

    // Creates a simple bitmap image using Aspose.Drawing and saves it to the given path.
    private static void CreateSampleImage(string path, int width, int height, Aspose.Drawing.Color background)
    {
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                graphics.Clear(background);
            }
            bitmap.Save(path);
        }
    }
}
