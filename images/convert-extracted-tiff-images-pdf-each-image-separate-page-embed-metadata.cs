using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace TiffToPdfConverter
{
    public class Converter
    {
        /// <summary>
        /// Converts all TIFF images in the specified folder to a single PDF.
        /// Each image is placed on its own page. Document metadata is also embedded.
        /// </summary>
        /// <param name="tiffFolder">Folder containing the source TIFF files.</param>
        /// <param name="outputPdfPath">Full path for the resulting PDF file.</param>
        public static void ConvertTiffFolderToPdf(string tiffFolder, string outputPdfPath)
        {
            // Validate inputs.
            if (string.IsNullOrEmpty(tiffFolder))
                throw new ArgumentException("Folder path must be provided.", nameof(tiffFolder));

            if (!Directory.Exists(tiffFolder))
                throw new DirectoryNotFoundException($"Folder not found: {tiffFolder}");

            // Gather all TIFF files (supports .tif and .tiff extensions).
            string[] tiffFiles = Directory.GetFiles(tiffFolder, "*.tif*");
            if (tiffFiles.Length == 0)
                throw new FileNotFoundException("No TIFF files found in the specified folder.", tiffFolder);

            // Create a new empty Word document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert each TIFF image on a separate page.
            for (int i = 0; i < tiffFiles.Length; i++)
            {
                // Insert the image.
                builder.InsertImage(tiffFiles[i]);

                // Add a page break after the image unless it is the last one.
                if (i < tiffFiles.Length - 1)
                    builder.InsertBreak(BreakType.PageBreak);
            }

            // Embed PDF metadata via built‑in document properties.
            doc.BuiltInDocumentProperties.Title = "Combined TIFF Images";
            doc.BuiltInDocumentProperties.Author = "Aspose.Words Converter";
            doc.BuiltInDocumentProperties.Subject = "TIFF to PDF conversion";
            doc.BuiltInDocumentProperties.Keywords = "TIFF, PDF, conversion, Aspose.Words";

            // Configure PDF save options (optional customizations can be added here).
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Example: ensure PDF 1.7 compliance.
                Compliance = PdfCompliance.Pdf17,

                // Example: embed the generator name (default is true, shown for clarity).
                ExportGeneratorName = true
            };

            // Save the assembled document as a PDF.
            doc.Save(outputPdfPath, pdfOptions);
        }

        // Example usage.
        public static void Main()
        {
            string sourceFolder = @"C:\Images\TiffFiles";
            string destinationPdf = @"C:\Output\CombinedImages.pdf";

            ConvertTiffFolderToPdf(sourceFolder, destinationPdf);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}
