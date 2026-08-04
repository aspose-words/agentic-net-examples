using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample Word document containing placeholders.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Report for _Customer_");
        builder.Writeln("Dear _Customer_,");
        builder.Writeln("Your order number is _OrderNumber_.");

        // Save the document as DOCX to simulate an existing input file.
        const string inputPath = "input.docx";
        source.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX file.
        Document doc = new Document(inputPath);

        // Perform find‑and‑replace operations.
        doc.Range.Replace("_Customer_", "Acme Corp");
        doc.Range.Replace("_OrderNumber_", "12345");

        // Export the modified document to Excel (XLSX) format.
        const string outputPath = "output.xlsx";
        doc.Save(outputPath, SaveFormat.Xlsx);

        // Verify that the Excel file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Expected output file '{outputPath}' was not created.");
    }
}
