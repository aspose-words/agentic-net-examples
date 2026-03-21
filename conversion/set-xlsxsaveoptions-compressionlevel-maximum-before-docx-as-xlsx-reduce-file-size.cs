using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Data 1");
        builder.InsertCell();
        builder.Write("Data 2");
        builder.EndTable();

        // Configure XlsxSaveOptions.
        XlsxSaveOptions saveOptions = new XlsxSaveOptions
        {
            CompressionLevel = CompressionLevel.Maximum,
            SaveFormat = SaveFormat.Xlsx
        };

        // Save the document as an XLSX file.
        doc.Save("Output.xlsx", saveOptions);
    }
}
