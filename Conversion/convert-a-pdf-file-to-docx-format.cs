using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

namespace PdfToDocxExample
{
    public class PdfToDocxConverter
    {
        /// <summary>
        /// Converts a PDF file to DOCX format.
        /// </summary>
        /// <param name="pdfPath">Full path to the source PDF file.</param>
        /// <param name="docxPath">Full path where the resulting DOCX file will be saved.</param>
        public static void Convert(string pdfPath, string docxPath)
        {
            // Load the PDF document using default PDF load options.
            var loadOptions = new PdfLoadOptions();
            var document = new Document(pdfPath, loadOptions);

            // Prepare save options for DOCX format.
            var saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);

            // Save the loaded document as DOCX.
            document.Save(docxPath, saveOptions);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Simple argument handling: first argument = input PDF, second argument = output DOCX.
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: PdfToDocxExample <input.pdf> <output.docx>");
                return;
            }

            string pdfPath = args[0];
            string docxPath = args[1];

            try
            {
                PdfToDocxConverter.Convert(pdfPath, docxPath);
                Console.WriteLine($"Successfully converted '{pdfPath}' to '{docxPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }
}
