using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Deterministic file names.
        const string sampleImagePath = "sample.png";
        const string odtPath = "sample.odt";
        const string headerFolder = "HeaderImages";
        const string footerFolder = "FooterImages";

        // --------------------------------------------------------------
        // 1. Create a sample image that will be inserted into header/footer.
        // --------------------------------------------------------------
        const int imgWidth = 100;
        const int imgHeight = 100;
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imgWidth, imgHeight))
        using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.White);
            bitmap.Save(sampleImagePath);
        }

        // --------------------------------------------------------------
        // 2. Build a document with header and footer containing the image.
        // --------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert image into the primary header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.InsertImage(sampleImagePath);

        // Insert image into the primary footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.InsertImage(sampleImagePath);

        // Save the document as ODT.
        doc.Save(odtPath, SaveFormat.Odt);

        // --------------------------------------------------------------
        // 3. Load the ODT document for image extraction.
        // --------------------------------------------------------------
        Document loadedDoc = new Document(odtPath);

        // Ensure output folders exist.
        Directory.CreateDirectory(headerFolder);
        Directory.CreateDirectory(footerFolder);

        int headerImageCount = 0;
        int footerImageCount = 0;

        // Iterate through each section's headers and footers.
        foreach (Section section in loadedDoc.Sections)
        {
            foreach (HeaderFooter hf in section.HeadersFooters)
            {
                bool isHeader = hf.HeaderFooterType == HeaderFooterType.HeaderPrimary ||
                                hf.HeaderFooterType == HeaderFooterType.HeaderFirst ||
                                hf.HeaderFooterType == HeaderFooterType.HeaderEven;

                // Collect all shape nodes that may contain images.
                NodeCollection shapes = hf.GetChildNodes(NodeType.Shape, true);
                foreach (Shape shape in shapes)
                {
                    if (shape.HasImage)
                    {
                        string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                        string fileName = $"image_{(isHeader ? ++headerImageCount : ++footerImageCount)}{extension}";
                        string folder = isHeader ? headerFolder : footerFolder;
                        string fullPath = Path.Combine(folder, fileName);

                        // Save the image to the appropriate folder.
                        shape.ImageData.Save(fullPath);
                    }
                }
            }
        }

        // --------------------------------------------------------------
        // 4. Validate that images were extracted.
        // --------------------------------------------------------------
        if (headerImageCount == 0)
            throw new InvalidOperationException("No images were extracted from headers.");
        if (footerImageCount == 0)
            throw new InvalidOperationException("No images were extracted from footers.");

        // Cleanup temporary files (optional).
        File.Delete(sampleImagePath);
        File.Delete(odtPath);
    }
}
