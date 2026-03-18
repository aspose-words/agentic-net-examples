using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace PdfBatchConversionExample
{
    public static class PdfBatchConverter
    {
        /// <summary>
        /// Converts all PDF files found in <paramref name="inputFolder"/> to HTML files in <paramref name="outputFolder"/>.
        /// The conversion preserves the original layout and embeds the fonts used in the source PDF.
        /// </summary>
        /// <param name="inputFolder">Folder that contains the source PDF files.</param>
        /// <param name="outputFolder">Folder where the resulting HTML files will be written.</param>
        public static void ConvertPdfsToHtml(string inputFolder, string outputFolder)
        {
            // Validate arguments.
            if (string.IsNullOrEmpty(inputFolder))
                throw new ArgumentException("Input folder path must be provided.", nameof(inputFolder));
            if (string.IsNullOrEmpty(outputFolder))
                throw new ArgumentException("Output folder path must be provided.", nameof(outputFolder));

            // Ensure the input folder exists.
            if (!Directory.Exists(inputFolder))
                throw new DirectoryNotFoundException($"Input folder not found: {inputFolder}");

            // Create the output folder if it does not exist.
            Directory.CreateDirectory(outputFolder);

            // Prepare the HtmlSaveOptions that will be used for every conversion.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportFontResources = true,
                ExportFontsAsBase64 = true,
                ExportImagesAsBase64 = true,
                ExportDocumentProperties = true,
                ExportHeadersFootersMode = ExportHeadersFootersMode.PerSection,
                CssSavingCallback = null,
                HtmlVersion = HtmlVersion.Xhtml,
                PrettyFormat = false
            };

            // Process each PDF file.
            foreach (string pdfPath in Directory.GetFiles(inputFolder, "*.pdf"))
            {
                // Load the PDF document.
                Document doc = new Document(pdfPath);

                // Build the output HTML file name (same base name, .html extension).
                string htmlFileName = Path.GetFileNameWithoutExtension(pdfPath) + ".html";
                string htmlPath = Path.Combine(outputFolder, htmlFileName);

                // Save the document as HTML using the prepared options.
                doc.Save(htmlPath, saveOptions);
            }
        }
    }

    class Program
    {
        /// <summary>
        /// Entry point required for a console application. Accepts two arguments:
        ///   1. Path to the folder containing PDF files.
        ///   2. Path to the folder where HTML files will be written.
        /// If arguments are missing, the program prints usage information.
        /// </summary>
        static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: PdfBatchConversionExample <inputFolder> <outputFolder>");
                return;
            }

            string inputFolder = args[0];
            string outputFolder = args[1];

            try
            {
                PdfBatchConverter.ConvertPdfsToHtml(inputFolder, outputFolder);
                Console.WriteLine($"Conversion completed. HTML files are in: {outputFolder}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
