using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentToExcelConversion
{
    class Program
    {
        static void Main()
        {
            // Load the source Word document. The constructor automatically detects the format.
            Document doc = new Document("SampleDocument.docx");

            // Create XlsxSaveOptions to specify Excel-specific saving behavior.
            XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
            {
                // Save each Word section as a separate worksheet.
                SectionMode = XlsxSectionMode.MultipleWorksheets,
                // Explicitly set the format to Xlsx (required by the options object).
                SaveFormat = SaveFormat.Xlsx
            };

            // Save the document as an Excel file using the configured options.
            doc.Save("SampleDocument.xlsx", xlsxOptions);
        }
    }
}
