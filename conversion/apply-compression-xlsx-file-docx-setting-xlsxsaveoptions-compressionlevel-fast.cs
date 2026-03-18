using System;
using Aspose.Words;
using Aspose.Words.Saving;

class XlsxCompressionExample
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("InputDocument.docx");

        // Create XlsxSaveOptions and configure fast compression.
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions();
        xlsxOptions.CompressionLevel = CompressionLevel.Fast;   // Fast (weaker) compression.
        xlsxOptions.SaveFormat = SaveFormat.Xlsx;               // Ensure the format is XLSX.

        // Save the document as an XLSX file using the specified options.
        doc.Save("CompressedOutput.xlsx", xlsxOptions);
    }
}
