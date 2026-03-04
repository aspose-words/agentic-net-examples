using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsExcelSaveExample
{
    class Program
    {
        static void Main()
        {
            // Load the source DOC document.
            Document doc = new Document("InputDocuments/SourceDocument.doc");

            // Create XlsxSaveOptions to control how the document is saved as an XLSX file.
            XlsxSaveOptions xlsxOptions = new XlsxSaveOptions();

            // Set the SectionMode to create a separate worksheet for each section.
            // This also determines the default worksheet names (e.g., "Section 1", "Section 2", ...).
            xlsxOptions.SectionMode = XlsxSectionMode.MultipleWorksheets;

            // Ensure the SaveFormat is explicitly set to Xlsx (optional, as the options class defaults to Xlsx).
            xlsxOptions.SaveFormat = SaveFormat.Xlsx;

            // Save the document as an XLSX file using the configured options.
            doc.Save("OutputDocuments/ConvertedDocument.xlsx", xlsxOptions);
        }
    }
}
