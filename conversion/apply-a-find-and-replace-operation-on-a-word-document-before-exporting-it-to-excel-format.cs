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
        builder.Writeln("Report for _CompanyName_ generated on _Date_.");

        // Save the source document as DOCX.
        const string inputPath = "input.docx";
        source.Save(inputPath, SaveFormat.Docx);

        // Load the document from the file.
        Document doc = new Document(inputPath);

        // Perform find‑and‑replace operations.
        doc.Range.Replace("_CompanyName_", "Acme Corp");
        doc.Range.Replace("_Date_", DateTime.Today.ToString("yyyy-MM-dd"));

        // Export the modified document to Excel format.
        const string outputPath = "output.xlsx";
        doc.Save(outputPath, SaveFormat.Xlsx);

        // Verify that the Excel file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Expected output Excel file was not created.");
    }
}
