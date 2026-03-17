using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsTableToPdf
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file that contains tables.
            string inputPath = @"C:\Docs\InputDocument.docx";

            // Path where the resulting PDF will be saved.
            string outputPath = @"C:\Docs\OutputDocument.pdf";

            // Load the DOCX document.
            Document doc = new Document(inputPath);

            // Convert any formatting defined in table styles to direct formatting.
            // This ensures that the visual appearance of tables is preserved
            // when the document is saved to a fixed‑page format such as PDF.
            doc.ExpandTableStylesToDirectFormatting();

            // Save the document as PDF. The Save method automatically determines
            // the format from the file extension, but we can also specify the
            // format explicitly using SaveFormat.Pdf.
            doc.Save(outputPath, SaveFormat.Pdf);
        }
    }
}
