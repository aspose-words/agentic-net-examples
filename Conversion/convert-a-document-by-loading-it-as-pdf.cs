using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

namespace AsposeWordsExamples
{
    public class PdfConverter
    {
        /// <summary>
        /// Loads a PDF file using <see cref="PdfLoadOptions"/> and saves it as a DOCX document.
        /// </summary>
        /// <param name="pdfPath">Full path to the source PDF file.</param>
        /// <param name="docxPath">Full path where the resulting DOCX file will be saved.</param>
        public void ConvertPdfToDocx(string pdfPath, string docxPath)
        {
            // Validate input arguments.
            if (string.IsNullOrEmpty(pdfPath))
                throw new ArgumentException("PDF path must be provided.", nameof(pdfPath));
            if (string.IsNullOrEmpty(docxPath))
                throw new ArgumentException("Output path must be provided.", nameof(docxPath));
            if (!File.Exists(pdfPath))
                throw new FileNotFoundException("PDF file not found.", pdfPath);

            // Create load options specific for PDF loading.
            PdfLoadOptions loadOptions = new PdfLoadOptions();

            // Load the PDF document into an Aspose.Words Document object.
            Document doc = new Document(pdfPath, loadOptions);

            // Save the loaded document in DOCX format.
            doc.Save(docxPath);
        }
    }

    class Program
    {
        /// <summary>
        /// Entry point of the console application.
        /// Usage: dotnet run <pdfPath> <docxPath>
        /// </summary>
        static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: <pdfPath> <docxPath>");
                return;
            }

            string pdfPath = args[0];
            string docxPath = args[1];

            try
            {
                var converter = new PdfConverter();
                converter.ConvertPdfToDocx(pdfPath, docxPath);
                Console.WriteLine($"PDF successfully converted to DOCX: {docxPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
