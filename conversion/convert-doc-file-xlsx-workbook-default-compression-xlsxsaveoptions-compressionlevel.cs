using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC file.
        Document doc = new Document("Input.doc");

        // Create XlsxSaveOptions with default compression (Normal).
        XlsxSaveOptions saveOptions = new XlsxSaveOptions();
        // Explicitly set the compression level to Normal (default).
        saveOptions.CompressionLevel = CompressionLevel.Normal;
        // Specify that the output format is XLSX.
        saveOptions.SaveFormat = SaveFormat.Xlsx;

        // Save the document as an XLSX workbook using the configured options.
        doc.Save("Output.xlsx", saveOptions);
    }
}
