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
        builder.Writeln("Sample content for XLSX conversion.");

        // Save the document as DOCX (bootstrap step).
        string docxPath = "sample.docx";
        doc.Save(docxPath, SaveFormat.Docx);

        // Load the DOCX file.
        Document loadedDoc = new Document(docxPath);

        // Configure XlsxSaveOptions with maximum compression.
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions();
        xlsxOptions.CompressionLevel = CompressionLevel.Maximum;
        xlsxOptions.SaveFormat = SaveFormat.Xlsx;

        // Save the document as XLSX using the configured options.
        string xlsxPath = "output.xlsx";
        loadedDoc.Save(xlsxPath, xlsxOptions);

        // Verify that the XLSX file was created.
        if (!File.Exists(xlsxPath))
            throw new InvalidOperationException("Expected XLSX output file was not created.");

        // Output the size of the compressed file.
        long fileSize = new FileInfo(xlsxPath).Length;
        Console.WriteLine($"XLSX saved with maximum compression. Size: {fileSize} bytes.");
    }
}
