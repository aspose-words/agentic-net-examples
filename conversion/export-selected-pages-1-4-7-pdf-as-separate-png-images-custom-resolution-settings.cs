using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsExport
{
    public class PdfToPngExporter
    {
        /// <summary>
        /// Exports pages 1, 4 and 7 (1‑based) of a PDF document to separate PNG files.
        /// Each image is rendered with a custom resolution (dots per inch).
        /// </summary>
        /// <param name="pdfPath">Full path to the source PDF file.</param>
        /// <param name="outputFolder">Folder where the PNG files will be written.</param>
        /// <param name="resolutionDpi">Resolution in DPI to use for rendering.</param>
        public static void ExportSelectedPages(string pdfPath, string outputFolder, float resolutionDpi)
        {
            // Load the PDF document.
            Document doc = new Document(pdfPath);

            // Prepare the image save options for PNG output.
            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                // Apply the custom resolution to both axes.
                Resolution = resolutionDpi
            };

            // Pages to export – 1‑based page numbers.
            int[] pagesToExport = { 1, 4, 7 };

            foreach (int pageNumber in pagesToExport)
            {
                // Convert to zero‑based index required by PageSet.
                int zeroBasedIndex = pageNumber - 1;

                // Guard against out‑of‑range page numbers.
                if (zeroBasedIndex < 0 || zeroBasedIndex >= doc.PageCount)
                    continue; // Skip invalid pages.

                // Set the PageSet to render only the current page.
                pngOptions.PageSet = new PageSet(zeroBasedIndex);

                // Build the output file name.
                string outFile = Path.Combine(outputFolder, $"Page_{pageNumber}.png");

                // Save the single page as a PNG image.
                doc.Save(outFile, pngOptions);
            }
        }

        // Example usage.
        public static void Main()
        {
            // Create a temporary PDF file with a few pages for demonstration.
            string tempDir = Path.Combine(Path.GetTempPath(), "AsposeDemo");
            Directory.CreateDirectory(tempDir);

            string pdfFile = Path.Combine(tempDir, "Sample.pdf");
            if (!File.Exists(pdfFile))
            {
                Document tempDoc = new Document();
                for (int i = 1; i <= 10; i++)
                {
                    Paragraph para = new Paragraph(tempDoc);
                    para.AppendChild(new Run(tempDoc, $"This is page {i}."));
                    tempDoc.FirstSection.Body.AppendChild(para);
                    if (i < 10)
                        tempDoc.AppendChild(new Section(tempDoc));
                }
                tempDoc.Save(pdfFile);
            }

            string outputDir = Path.Combine(tempDir, "ExportedPages");
            Directory.CreateDirectory(outputDir);

            float dpi = 300f; // Custom resolution.

            ExportSelectedPages(pdfFile, outputDir, dpi);

            Console.WriteLine($"Export completed. PNG files are located in: {outputDir}");
        }
    }
}
