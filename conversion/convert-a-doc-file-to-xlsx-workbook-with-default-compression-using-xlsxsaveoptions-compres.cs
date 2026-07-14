using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOC file.
        const string inputPath = "sample.doc";
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("This is a sample DOC file.");
        source.Save(inputPath, SaveFormat.Doc);

        // Load the DOC file.
        Document doc = new Document(inputPath);

        // Set up XlsxSaveOptions with the default compression level.
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            CompressionLevel = CompressionLevel.Normal,
            SaveFormat = SaveFormat.Xlsx
        };

        // Convert and save as XLSX.
        const string outputPath = "output.xlsx";
        doc.Save(outputPath, xlsxOptions);

        // Verify that the XLSX file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Expected output XLSX was not created.");
    }
}
