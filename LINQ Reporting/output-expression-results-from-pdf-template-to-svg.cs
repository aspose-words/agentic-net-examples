using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace OutputExpressionResultsFromPdfTemplateToSvg
{
    class Program
    {
        /// <summary>
        /// Entry point of the example.
        /// </summary>
        static void Main(string[] args)
        {
            // Example usage – adjust the paths as needed.
            string pdfPath = @"C:\Path\To\Your\Template.pdf";
            string outputFolder = @"C:\Path\To\Output\Svg";

            PdfToSvgConverter.Convert(pdfPath, outputFolder);

            Console.WriteLine("Conversion completed. SVG files are saved to: " + outputFolder);
        }
    }

    /// <summary>
    /// Loads a PDF document, evaluates all field expressions, and saves the result as SVG.
    /// </summary>
    class PdfToSvgConverter
    {
        public static void Convert(string pdfPath, string outputFolder)
        {
            // Load the PDF document into an Aspose.Words Document object.
            Document doc = new Document(pdfPath);

            // Evaluate all fields (expressions) so that calculated results are present in the output.
            doc.UpdateFields();

            // Ensure the output directory exists.
            Directory.CreateDirectory(outputFolder);

            // Configure SVG save options.
            SvgSaveOptions options = new SvgSaveOptions
            {
                // Render text as placed glyphs (curves) to avoid font‑dependency issues.
                TextOutputMode = SvgTextOutputMode.UsePlacedGlyphs,
                // Do not add a page border around the SVG content.
                ShowPageBorder = false,
                // Make the SVG fill the available viewport (width/height = 100%).
                FitToViewPort = true
            };

            // Save the document as SVG. For multi‑page documents Aspose.Words creates separate SVG files
            // (Result.svg, Result.001.svg, Result.002.svg, …) in the specified folder.
            string outputPath = Path.Combine(outputFolder, "Result.svg");
            doc.Save(outputPath, options);
        }
    }
}
