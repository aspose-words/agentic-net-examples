using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOC file.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample DOC content.");
        source.Save("input.doc", SaveFormat.Doc);

        // Load the DOC file.
        Document doc = new Document("input.doc");

        // Prepare XLSX save options with default compression (Normal).
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            CompressionLevel = CompressionLevel.Normal,
            SaveFormat = SaveFormat.Xlsx
        };

        // Convert and save as XLSX.
        doc.Save("output.xlsx", xlsxOptions);

        // Verify that the XLSX file was created.
        if (!File.Exists("output.xlsx"))
            throw new InvalidOperationException("The expected XLSX output file was not created.");
    }
}
