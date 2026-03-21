using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractBackgroundImages
{
    static void Main()
    {
        // Use a path relative to the executable so the file always exists.
        string baseDir = AppContext.BaseDirectory;
        string inputPath = Path.Combine(baseDir, "input.docx");
        string outputDir = Path.Combine(baseDir, "BackgroundImages");

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputDir);

        // If the input file does not exist, create a minimal DOCX file.
        if (!File.Exists(inputPath))
        {
            var emptyDoc = new Document();
            emptyDoc.Save(inputPath);
        }

        // Load the document.
        Document doc = new Document(inputPath);

        // -----------------------------------------------------------------
        // 1. Extract the document‑level background shape (if any).
        // -----------------------------------------------------------------
        Shape bgShape = doc.BackgroundShape;
        if (bgShape != null && bgShape.HasImage)
        {
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

        // Indicate completion.
        Console.WriteLine($"Extraction complete. Images saved to: {outputDir}");
    }
}
