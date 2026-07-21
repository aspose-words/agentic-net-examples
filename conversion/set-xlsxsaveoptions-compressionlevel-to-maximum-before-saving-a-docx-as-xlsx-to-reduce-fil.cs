using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample content for conversion.");

        string inputPath = "sample.docx";
        doc.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX document.
        Document loadedDoc = new Document(inputPath);

        // Configure XlsxSaveOptions with maximum compression.
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            CompressionLevel = CompressionLevel.Maximum,
            SaveFormat = SaveFormat.Xlsx
        };

        string outputPath = "converted.xlsx";
        loadedDoc.Save(outputPath, xlsxOptions);

        // Verify that the XLSX file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The XLSX output file was not created.");

        // Optional: clean up the temporary DOCX file.
        if (File.Exists(inputPath))
            File.Delete(inputPath);
    }
}
