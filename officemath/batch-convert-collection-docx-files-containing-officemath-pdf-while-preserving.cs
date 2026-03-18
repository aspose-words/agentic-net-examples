using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

namespace AsposeWordsBatchConversion
{
    /// <summary>
    /// Provides functionality to batch‑convert DOCX files that contain OfficeMath equations to PDF.
    /// The conversion preserves equation fidelity by loading shapes as OfficeMath objects and by using
    /// high‑quality PDF save options.
    /// </summary>
    public static class DocxToPdfConverter
    {
        /// <summary>
        /// Converts each DOCX file in <paramref name="inputFiles"/> to a PDF file placed in <paramref name="outputFolder"/>.
        /// </summary>
        /// <param name="inputFiles">Full paths of the source DOCX files.</param>
        /// <param name="outputFolder">Folder where the resulting PDF files will be saved.</param>
        public static void ConvertDocsToPdf(string[] inputFiles, string outputFolder)
        {
            // Ensure the output directory exists.
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Configure PDF save options once – they will be reused for every document.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Embed all fonts to guarantee that equations render exactly as in the source.
                EmbedFullFonts = true,

                // Use the highest rendering quality (slower but more accurate).
                UseHighQualityRendering = true,

                // Optional: set compliance to PDF 1.7 (default) – can be changed if needed.
                Compliance = PdfCompliance.Pdf17,

                // Ensure fields (including equation fields) are up‑to‑date before saving.
                UpdateFields = true
            };

            // Iterate over each input file.
            foreach (string inputPath in inputFiles)
            {
                // Validate the source file.
                if (string.IsNullOrWhiteSpace(inputPath) || !File.Exists(inputPath))
                    continue; // Skip missing or invalid entries.

                // LoadOptions: convert any shape that contains EquationXML to an OfficeMath object.
                LoadOptions loadOptions = new LoadOptions
                {
                    ConvertShapeToOfficeMath = true
                };

                // Load the DOCX document using the load options.
                Document doc = new Document(inputPath, loadOptions);

                // Build the output PDF file name (same base name, .pdf extension).
                string outputFileName = Path.GetFileNameWithoutExtension(inputPath) + ".pdf";
                string outputPath = Path.Combine(outputFolder, outputFileName);

                // Save the document as PDF using the prepared options.
                doc.Save(outputPath, pdfOptions);
            }
        }
    }

    /// <summary>
    /// Console entry point used by the build system. Demonstrates how to call the batch converter.
    /// </summary>
    public static class Program
    {
        /// <summary>
        /// Expected arguments:
        ///   1. Output folder path.
        ///   2..N. One or more input DOCX file paths.
        /// </summary>
        public static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: dotnet run <outputFolder> <inputFile1.docx> [<inputFile2.docx> ...]");
                return;
            }

            string outputFolder = args[0];
            string[] inputFiles = new string[args.Length - 1];
            Array.Copy(args, 1, inputFiles, 0, inputFiles.Length);

            try
            {
                DocxToPdfConverter.ConvertDocsToPdf(inputFiles, outputFolder);
                Console.WriteLine($"Conversion completed. PDFs saved to '{outputFolder}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }
}
