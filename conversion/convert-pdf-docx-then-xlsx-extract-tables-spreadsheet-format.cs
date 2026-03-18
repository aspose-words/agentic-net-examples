using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsConversion
{
    public static class PdfToXlsxConverter
    {
        /// <summary>
        /// Converts a PDF file to DOCX and then to XLSX.
        /// The resulting XLSX file contains each table from the original document
        /// on a separate worksheet.
        /// </summary>
        /// <param name="pdfPath">Full path to the source PDF file.</param>
        /// <param name="docxPath">Full path where the intermediate DOCX will be saved.</param>
        /// <param name="xlsxPath">Full path where the final XLSX will be saved.</param>
        public static void Convert(string pdfPath, string docxPath, string xlsxPath)
        {
            // Load the PDF document.
            Document pdfDocument = new Document(pdfPath);

            // Save the loaded PDF as a DOCX file.
            pdfDocument.Save(docxPath, SaveFormat.Docx);

            // Load the newly created DOCX document.
            Document docxDocument = new Document(docxPath);

            // Configure XLSX save options to place each section (tables) on its own worksheet.
            XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
            {
                SectionMode = XlsxSectionMode.MultipleWorksheets,
                SaveFormat = SaveFormat.Xlsx
            };

            // Save the DOCX as an XLSX spreadsheet.
            docxDocument.Save(xlsxPath, xlsxOptions);
        }
    }

    class Program
    {
        /// <summary>
        /// Entry point required for a console application.
        /// Adjust the file paths below to point to your actual PDF, DOCX, and XLSX locations.
        /// </summary>
        static void Main(string[] args)
        {
            // Example file locations – replace with real paths or pass via command‑line arguments.
            string pdfPath = "input.pdf";
            string docxPath = "intermediate.docx";
            string xlsxPath = "output.xlsx";

            // If arguments are supplied, use them (pdf, docx, xlsx).
            if (args.Length >= 3)
            {
                pdfPath = args[0];
                docxPath = args[1];
                xlsxPath = args[2];
            }

            try
            {
                PdfToXlsxConverter.Convert(pdfPath, docxPath, xlsxPath);
                Console.WriteLine($"Conversion succeeded. XLSX saved to '{xlsxPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Conversion failed: {ex.Message}");
            }
        }
    }
}
