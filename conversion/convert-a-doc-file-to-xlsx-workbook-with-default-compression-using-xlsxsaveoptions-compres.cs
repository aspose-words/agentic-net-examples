using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOC file.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample DOC content for conversion to XLSX.");
        sourceDoc.Save("input.doc", SaveFormat.Doc);

        // Load the DOC file.
        Document doc = new Document("input.doc");

        // Configure XLSX save options with a valid compression level.
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            CompressionLevel = CompressionLevel.Normal // Use a defined compression level.
        };

        // Save the document as XLSX.
        doc.Save("output.xlsx", xlsxOptions);

        // Verify that the XLSX file was created.
        if (!File.Exists("output.xlsx"))
        {
            throw new InvalidOperationException("Expected output XLSX was not created.");
        }
    }
}
