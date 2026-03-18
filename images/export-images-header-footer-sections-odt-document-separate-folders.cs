using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing; // <-- added for Shape
using Aspose.Words.Saving;

class ExportHeaderFooterImages
{
    static void Main()
    {
        // Path to the source ODT document.
        string sourcePath = @"C:\Docs\SourceDocument.odt";

        // Folders where header and footer images will be saved.
        string headerImagesFolder = Path.Combine(@"C:\Docs\HeaderImages");
        string footerImagesFolder = Path.Combine(@"C:\Docs\FooterImages");

        // Ensure the output folders exist.
        Directory.CreateDirectory(headerImagesFolder);
        Directory.CreateDirectory(footerImagesFolder);

        // Load the ODT document.
        Document doc = new Document(sourcePath);

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
        string copyPath = Path.Combine(@"C:\Docs", "SourceDocument_Copy.odt");
        doc.Save(copyPath);
    }
}
