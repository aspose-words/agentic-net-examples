using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create deterministic sample images for header and footer.
        const string headerImagePath = "header.png";
        const string footerImagePath = "footer.png";

        CreateSampleImage(headerImagePath, 200, 50, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage(footerImagePath, 200, 50, Aspose.Drawing.Color.LightGreen);

        // Create a new document and insert the images into header and footer.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert image into the primary header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.InsertImage(headerImagePath);

        // Insert image into the primary footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.InsertImage(footerImagePath);

        // Save the document as ODT.
        const string odtPath = "sample.odt";
        doc.Save(odtPath, SaveFormat.Odt);

        // Load the saved ODT document.
        Document loadedDoc = new Document(odtPath);

        // Prepare output folders.
        const string headerFolder = "HeaderImages";
        const string footerFolder = "FooterImages";
        Directory.CreateDirectory(headerFolder);
        Directory.CreateDirectory(footerFolder);

        // Extract images from headers and footers.
        foreach (Section section in loadedDoc.Sections)
        {
            foreach (HeaderFooter hf in section.HeadersFooters)
            {
                // Determine target folder based on header/footer type.
                string targetFolder = (hf.HeaderFooterType == HeaderFooterType.HeaderPrimary ||
                                       hf.HeaderFooterType == HeaderFooterType.HeaderFirst ||
                                       hf.HeaderFooterType == HeaderFooterType.HeaderEven)
                                      ? headerFolder
                                      : footerFolder;

                // Find all Shape nodes inside the header/footer.
                NodeCollection shapes = hf.GetChildNodes(NodeType.Shape, true);
                int imageIndex = 0;
                foreach (Shape shape in shapes)
                {
                    if (shape.HasImage)
                    {
                        string imageFileName = Path.Combine(targetFolder,
                            $"image-{hf.HeaderFooterType}-{imageIndex}.png");
                        shape.ImageData.Save(imageFileName);
                        imageIndex++;
                    }
                }
            }
        }

        // Validate that images were extracted.
        ValidateExtraction(headerFolder, "header");
        ValidateExtraction(footerFolder, "footer");
    }

    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(backColor);
                // Simple deterministic drawing: a filled rectangle.
                using (SolidBrush brush = new SolidBrush(Aspose.Drawing.Color.DarkGray))
                {
                    graphics.FillRectangle(brush, width / 4, height / 4, width / 2, height / 2);
                }
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }

        // Ensure the file exists.
        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create image file: {filePath}");
    }

    private static void ValidateExtraction(string folderPath, string description)
    {
        string[] files = Directory.GetFiles(folderPath);
        if (files.Length == 0)
            throw new InvalidOperationException($"No images were extracted for {description} into folder '{folderPath}'.");
    }
}
