using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractBackgroundImages
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\input.docx";

        // Folder where extracted PNG images will be saved.
        string outputDir = @"C:\Docs\BackgroundImages";
        Directory.CreateDirectory(outputDir);

        // Load the document.
        Document doc = new Document(inputPath);

        // -----------------------------------------------------------------
        // 1. Extract the document‑level background shape (if any).
        // -----------------------------------------------------------------
        Shape bgShape = doc.BackgroundShape;
        if (bgShape != null && bgShape.HasImage)
        {
            // Save the image directly – no System.Drawing required.
            string outPath = Path.Combine(outputDir, "DocumentBackground.png");
            bgShape.ImageData.Save(outPath);
        }

        // -----------------------------------------------------------------
        // 2. Extract background images that may be stored in headers/footers.
        // -----------------------------------------------------------------
        int imageIndex = 0;
        foreach (Section section in doc.Sections)
        {
            foreach (HeaderFooter hf in section.HeadersFooters)
            {
                // Get all shapes inside the header/footer.
                NodeCollection shapes = hf.GetChildNodes(NodeType.Shape, true);
                foreach (Shape shape in shapes)
                {
                    if (shape.HasImage)
                    {
                        string outPath = Path.Combine(outputDir,
                            $"HeaderFooterBackground_{imageIndex++}.png");
                        shape.ImageData.Save(outPath);
                    }
                }
            }
        }

        // The images are now saved as PNG files in the specified folder.
    }
}
