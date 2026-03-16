using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class ExtractImagesBySection
{
    static void Main()
    {
        // Path to the source Word document.
        string inputPath = @"C:\Docs\InputDocument.docx";

        // Directory where extracted images will be saved.
        string outputDir = @"C:\Docs\ExtractedImages";

        // Ensure the output directory exists.
        if (!Directory.Exists(outputDir))
            Directory.CreateDirectory(outputDir);

        // Load the document (lifecycle rule: use the Document constructor for loading).
        Document doc = new Document(inputPath);

        // Iterate through each section in the document.
        for (int secIdx = 0; secIdx < doc.Sections.Count; secIdx++)
        {
            // Get the current section.
            Section section = doc.Sections[secIdx];

            // Collect all Shape nodes within the section (including nested shapes).
            NodeCollection shapes = section.GetChildNodes(NodeType.Shape, true);

            int imageIdx = 0;

            // Process each shape that contains an image.
            foreach (Shape shape in shapes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    // Build a unique file name that includes the section index and image index.
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string fileName = $"Section_{secIdx + 1}_Image_{++imageIdx}{extension}";
                    string fullPath = Path.Combine(outputDir, fileName);

                    // Save the image data to the file system (lifecycle rule: use ImageData.Save).
                    shape.ImageData.Save(fullPath);
                }
            }
        }

        // Optionally, save the (unchanged) document to a new file if needed.
        // This demonstrates the save lifecycle rule.
        string outputDocPath = Path.Combine(outputDir, "ProcessedDocument.docx");
        doc.Save(outputDocPath);
    }
}
