using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace ReportGeneratorApp
{
    public class ReportGenerator
    {
        /// <summary>
        /// Creates a new blank Word document, adds a "Confidential" text watermark,
        /// and saves it to the specified file path.
        /// </summary>
        /// <param name="outputPath">Full file name (including extension) where the document will be saved.</param>
        public void GenerateReport(string outputPath)
        {
            // Create a new blank document.
            Document doc = new Document();

            // Configure watermark appearance.
            TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
            {
                FontFamily = "Arial",
                FontSize = 36,
                Color = Color.Red,
                Layout = WatermarkLayout.Diagonal,
                IsSemitrasparent = false // Fully opaque.
            };

            // Apply the text watermark to the document.
            doc.Watermark.SetText("Confidential", watermarkOptions);

            // Save the document to the desired location.
            doc.Save(outputPath);
        }
    }

    class Program
    {
        /// <summary>
        /// Entry point required for a console application.
        /// </summary>
        static void Main(string[] args)
        {
            // Simple argument handling – first argument is the output file path.
            string outputPath = args.Length > 0 ? args[0] : "ConfidentialReport.docx";

            var generator = new ReportGenerator();
            generator.GenerateReport(outputPath);

            Console.WriteLine($"Report generated with confidential watermark at: {outputPath}");
        }
    }
}
