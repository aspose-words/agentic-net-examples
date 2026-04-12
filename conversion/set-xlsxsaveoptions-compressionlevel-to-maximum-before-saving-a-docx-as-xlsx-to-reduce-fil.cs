using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string inputDocxPath = "Sample.docx";
        string outputXlsxPath = "Compressed.xlsx";

        // -----------------------------------------------------------------
        // Create a sample DOCX document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose.Words!");
        // Save the document as DOCX to simulate an existing source file.
        doc.Save(inputDocxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Load the DOCX document from disk.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputDocxPath);

        // -----------------------------------------------------------------
        // Configure XlsxSaveOptions with maximum compression.
        // -----------------------------------------------------------------
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            CompressionLevel = CompressionLevel.Maximum,
            SaveFormat = SaveFormat.Xlsx
        };

        // Save the loaded document as XLSX using the configured options.
        loadedDoc.Save(outputXlsxPath, xlsxOptions);

        // -----------------------------------------------------------------
        // Validate that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputXlsxPath))
        {
            throw new FileNotFoundException($"The file '{outputXlsxPath}' was not created.");
        }

        // Optional: display the size of the compressed file.
        FileInfo info = new FileInfo(outputXlsxPath);
        Console.WriteLine($"Compressed XLSX saved successfully. Size: {info.Length} bytes.");
    }
}
