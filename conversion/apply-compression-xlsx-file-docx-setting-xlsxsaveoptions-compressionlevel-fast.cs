using System;
using Aspose.Words;
using Aspose.Words.Saving;

class XlsxCompressionExample
{
    static void Main()
    {
        // Create a new document and add some content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words!");

        // Configure XlsxSaveOptions for fast compression.
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            CompressionLevel = CompressionLevel.Fast, // Fast (weaker) compression.
            SaveFormat = SaveFormat.Xlsx               // Ensure the format is XLSX.
        };

        // Save the document as an XLSX file using the specified options.
        doc.Save("CompressedOutput.xlsx", xlsxOptions);
    }
}
