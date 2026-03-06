using System;
using Aspose.Words;

namespace AsposeWordsConversion
{
    /// <summary>
    /// Provides functionality to convert DOCX files to PDF using Aspose.Words.
    /// </summary>
    public static class DocxToPdfConverter
    {
        /// <summary>
        /// Converts a DOCX document to a PDF file.
        /// </summary>
        /// <param name="docxPath">Full path to the source DOCX file.</param>
        /// <param name="pdfPath">Full path where the resulting PDF will be saved.</param>
        public static void Convert(string docxPath, string pdfPath)
        {
            // Load the DOCX document from the specified file.
            Document doc = new Document(docxPath);

            // Save the document as PDF using the overload that specifies the format.
            // This follows the documented Save(string, SaveFormat) rule.
            doc.Save(pdfPath, SaveFormat.Pdf);
        }

        // Example usage.
        public static void Main()
        {
            // Adjust these paths as needed.
            string sourceDocx = @"C:\Input\SampleDocument.docx";
            string targetPdf = @"C:\Output\SampleDocument.pdf";

            Convert(sourceDocx, targetPdf);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}
