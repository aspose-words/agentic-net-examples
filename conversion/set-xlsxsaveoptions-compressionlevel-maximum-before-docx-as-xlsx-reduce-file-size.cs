using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("Input.docx");

        // Create XlsxSaveOptions using the default constructor.
        XlsxSaveOptions xlsxSaveOptions = new XlsxSaveOptions();

        // Set the compression level to Maximum to reduce the XLSX file size.
        xlsxSaveOptions.CompressionLevel = CompressionLevel.Maximum;

        // Explicitly specify the target format (XLSX).
        xlsxSaveOptions.SaveFormat = SaveFormat.Xlsx;

        // Save the document as an XLSX file using the configured save options.
        doc.Save("Output.xlsx", xlsxSaveOptions);
    }
}
