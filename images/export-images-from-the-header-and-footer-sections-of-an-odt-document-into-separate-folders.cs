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
        // Define file and folder names
        const string docPath = "Sample.odt";
        const string headerImagePath = "header.png";
        const string footerImagePath = "footer.png";
        const string headerFolder = "HeaderImages";
        const string footerFolder = "FooterImages";

        // Ensure clean environment
        foreach (var path in new[] { docPath, headerImagePath, footerImagePath })
            if (File.Exists(path)) File.Delete(path);
        foreach (var folder in new[] { headerFolder, footerFolder })
            if (Directory.Exists(folder)) Directory.Delete(folder, true);

        // -------------------------------------------------
        // 1. Create sample images using Aspose.Drawing
        // -------------------------------------------------
        CreateSampleImage(headerImagePath, 200, 50, Aspose.Drawing.Color.LightBlue, "Header");
        CreateSampleImage(footerImagePath, 200, 50, Aspose.Drawing.Color.LightGreen, "Footer");

        // -------------------------------------------------
        // 2. Create an ODT document and insert images into header and footer
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert image into primary header
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.InsertImage(headerImagePath);
        builder.Writeln(); // ensure the header has some text after the image

        // Insert image into primary footer
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.InsertImage(footerImagePath);
        builder.Writeln(); // ensure the footer has some text after the image

        // Save the document as ODT
        doc.Save(docPath, SaveFormat.Odt);

        // -------------------------------------------------
        // 3. Load the document (demonstrating load lifecycle)
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // -------------------------------------------------
        // 4. Extract images from headers and footers into separate folders
        // -------------------------------------------------
        Directory.CreateDirectory(headerFolder);
        Directory.CreateDirectory(footerFolder);

        int headerImageIndex = 0;
        int footerImageIndex = 0;

        foreach (Section section in loadedDoc.Sections)
        {
            // Process header
            HeaderFooter header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
            if (header != null)
            {
                foreach (Shape shape in header.GetChildNodes(NodeType.Shape, true).OfType<Shape>())
                {
                    if (shape.HasImage)
                    {
                        string ext = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                        string fileName = Path.Combine(headerFolder,
                            $"header_image_{headerImageIndex}{ext}");
                        shape.ImageData.Save(fileName);
                        headerImageIndex++;
                    }
                }
            }

            // Process footer
            HeaderFooter footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
            if (footer != null)
            {
                foreach (Shape shape in footer.GetChildNodes(NodeType.Shape, true).OfType<Shape>())
                {
                    if (shape.HasImage)
                    {
                        string ext = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                        string fileName = Path.Combine(footerFolder,
                            $"footer_image_{footerImageIndex}{ext}");
                        shape.ImageData.Save(fileName);
                        footerImageIndex++;
                    }
                }
            }
        }

        // -------------------------------------------------
        // 5. Validation – ensure at least one image was saved per folder
        // -------------------------------------------------
        if (!Directory.EnumerateFiles(headerFolder).Any())
            throw new InvalidOperationException("No header images were extracted.");
        if (!Directory.EnumerateFiles(footerFolder).Any())
            throw new InvalidOperationException("No footer images were extracted.");

        Console.WriteLine("Header and footer images have been exported successfully.");
    }

    // Helper method to create a deterministic sample PNG image
    private static void CreateSampleImage(string filePath, int width, int height,
        Aspose.Drawing.Color backgroundColor, string text)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(backgroundColor);
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 20))
                {
                    graphics.DrawString(text, font, Aspose.Drawing.Brushes.Black, new PointF(10, 10));
                }
            }

            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}
