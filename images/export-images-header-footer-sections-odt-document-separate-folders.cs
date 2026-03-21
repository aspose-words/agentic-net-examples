using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExportHeaderFooterImages
{
    static void Main()
    {
        // Path to the source ODT document (use a file in the current directory if it exists,
        // otherwise create a new empty document).
        string sourcePath = Path.Combine(Directory.GetCurrentDirectory(), "SourceDocument.odt");

        Document doc;
        if (File.Exists(sourcePath))
        {
            // Load the existing ODT document.
            doc = new Document(sourcePath);
        }
        else
        {
            // Create a new empty document and add a simple header/footer so the example can run.
            doc = new Document();
            Section section = doc.Sections[0];

            // Add a header with a placeholder shape (no image).
            HeaderFooter header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
            section.HeadersFooters.Add(header);
            Shape headerShape = new Shape(doc, ShapeType.Image);
            Paragraph headerPara = new Paragraph(doc);
            headerPara.AppendChild(headerShape);
            header.AppendChild(headerPara);

            // Add a footer with a placeholder shape.
            HeaderFooter footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
            section.HeadersFooters.Add(footer);
            Shape footerShape = new Shape(doc, ShapeType.Image);
            Paragraph footerPara = new Paragraph(doc);
            footerPara.AppendChild(footerShape);
            footer.AppendChild(footerPara);
        }

        // Folders where header and footer images will be saved (relative to the current directory).
        string headerImagesFolder = Path.Combine(Directory.GetCurrentDirectory(), "HeaderImages");
        string footerImagesFolder = Path.Combine(Directory.GetCurrentDirectory(), "FooterImages");

        // Ensure the output folders exist.
        Directory.CreateDirectory(headerImagesFolder);
        Directory.CreateDirectory(footerImagesFolder);

        // Counters to generate unique file names.
        int headerImageIndex = 0;
        int footerImageIndex = 0;

        // Iterate through every section of the document.
        foreach (Section section in doc.Sections)
        {
            // Iterate through each header/footer in the current section.
            foreach (HeaderFooter headerFooter in section.HeadersFooters)
            {
                bool isHeader = headerFooter.IsHeader;

                // Get all Shape nodes (including those inside tables) that may contain images.
                NodeCollection shapeNodes = headerFooter.GetChildNodes(NodeType.Shape, true);

                foreach (Shape shape in shapeNodes.OfType<Shape>())
                {
                    // Process only shapes that actually contain an image.
                    if (!shape.HasImage)
                        continue;

                    // Determine the appropriate file extension for the image type.
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                    // Build a unique file name.
                    string fileName = (isHeader
                        ? $"header_{headerImageIndex++}"
                        : $"footer_{footerImageIndex++}") + extension;

                    // Choose the correct output folder.
                    string targetFolder = isHeader ? headerImagesFolder : footerImagesFolder;

                    // Save the image to the file system.
                    shape.ImageData.Save(Path.Combine(targetFolder, fileName));
                }
            }
        }

        // Optional: save a copy of the document to demonstrate the save rule.
        string copyPath = Path.Combine(Directory.GetCurrentDirectory(), "SourceDocument_Copy.odt");
        doc.Save(copyPath);
    }
}
