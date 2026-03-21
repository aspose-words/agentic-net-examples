using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX that contains embedded spreadsheet charts.
        // Use a relative path so the example can run without requiring a specific absolute location.
        string inputPath = Path.Combine(AppContext.BaseDirectory, "ChartDocument.docx");

        // Folder where extracted PNG images will be saved.
        string outputFolder = Path.Combine(AppContext.BaseDirectory, "ExtractedCharts");

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Verify that the input file exists before attempting to load it.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            Console.WriteLine("Place a DOCX file named 'ChartDocument.docx' in the application directory and rerun the program.");
            return;
        }

        // Load the Word document.
        Document doc = new Document(inputPath);

        // Retrieve all Shape nodes from the document (including those inside headers/footers).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int chartIndex = 0;

        // Iterate through each shape and look for OLE objects that represent Excel charts.
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // An embedded chart is stored as an OLE object with a ProgId that contains "Chart".
            if (shape.OleFormat != null &&
                !string.IsNullOrEmpty(shape.OleFormat.ProgId) &&
                shape.OleFormat.ProgId.IndexOf("Chart", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // Build the output file name.
                string outFile = Path.Combine(outputFolder, $"Chart_{chartIndex}.png");

                // Configure image save options – PNG format.
                ImageSaveOptions imgOptions = new ImageSaveOptions(SaveFormat.Png);

                // Render the shape (the chart) to an image file.
                shape.GetShapeRenderer().Save(outFile, imgOptions);

                chartIndex++;
            }
        }

        Console.WriteLine($"Extracted {chartIndex} chart image(s) to \"{outputFolder}\".");
    }
}
