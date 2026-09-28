using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare sample images for header and footer.
        const string headerImagePath = "header.png";
        const string footerImagePath = "footer.png";

        CreateSampleImage(headerImagePath, 200, 100, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage(footerImagePath, 200, 100, Aspose.Drawing.Color.LightGreen);

        // Create a sample ODT document with header and footer containing the images.
        const string sourceDocPath = "sample.odt";
        CreateDocumentWithHeaderFooterImages(sourceDocPath, headerImagePath, footerImagePath);

        // Load the document for extraction.
        Document doc = new Document(sourceDocPath);

        // Output folders.
        const string headerOutputFolder = "HeaderImages";
        const string footerOutputFolder = "FooterImages";
        Directory.CreateDirectory(headerOutputFolder);
        Directory.CreateDirectory(footerOutputFolder);

        int headerImageCount = 0;
        int footerImageCount = 0;

        // Iterate through sections and their header/footer collections.
        foreach (Section section in doc.Sections)
        {
            foreach (HeaderFooter hf in section.HeadersFooters)
            {
                // Determine target folder based on header/footer type.
                string targetFolder;
                bool isHeader = hf.HeaderFooterType == HeaderFooterType.HeaderPrimary ||
                                hf.HeaderFooterType == HeaderFooterType.HeaderFirst ||
                                hf.HeaderFooterType == HeaderFooterType.HeaderEven;
                if (isHeader)
                {
                    targetFolder = headerOutputFolder;
                }
                else
                {
                    targetFolder = footerOutputFolder;
                }

                // Find all Shape nodes that contain images.
                NodeCollection shapes = hf.GetChildNodes(NodeType.Shape, true);
                foreach (Shape shape in shapes)
                {
                    if (shape.HasImage)
                    {
                        string fileName = $"image-{Guid.NewGuid()}.png";
                        string fullPath = Path.Combine(targetFolder, fileName);
                        shape.ImageData.Save(fullPath);
                        if (isHeader)
                            headerImageCount++;
                        else
                            footerImageCount++;
                    }
                }
            }
        }

        // Validation.
        if (headerImageCount == 0)
            throw new InvalidOperationException("No header images were extracted.");
        if (footerImageCount == 0)
            throw new InvalidOperationException("No footer images were extracted.");

        Console.WriteLine($"Extraction complete. Header images: {headerImageCount}, Footer images: {footerImageCount}");
    }

    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(backColor);
            }
            bitmap.Save(filePath);
        }
    }

    private static void CreateDocumentWithHeaderFooterImages(string docPath, string headerImgPath, string footerImgPath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert header image.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.InsertImage(headerImgPath);

        // Insert footer image.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.InsertImage(footerImgPath);

        // Add a simple body paragraph.
        builder.MoveToDocumentEnd();
        builder.Writeln("Sample document body text.");

        doc.Save(docPath, SaveFormat.Odt);
    }
}
