using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample document with a table and save it as PDF (input.pdf).
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        // Insert a simple 2x2 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Row 1, Cell 1");
        builder.InsertCell();
        builder.Writeln("Row 1, Cell 2");
        builder.EndTable();

        string pdfPath = "input.pdf";
        sampleDoc.Save(pdfPath, SaveFormat.Pdf);

        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
            throw new InvalidOperationException("PDF file was not created successfully.");

        // Step 2: Load the PDF and convert it to DOCX.
        LoadOptions pdfLoadOptions = new LoadOptions();
        pdfLoadOptions.LoadFormat = LoadFormat.Pdf; // Specify that the source is a PDF.
        Document pdfDocument = new Document(pdfPath, pdfLoadOptions);

        string docxPath = "output.docx";
        pdfDocument.Save(docxPath, SaveFormat.Docx);

        if (!File.Exists(docxPath) || new FileInfo(docxPath).Length == 0)
            throw new InvalidOperationException("DOCX file was not created successfully.");

        // Step 3: Load the DOCX (containing the table) and convert it to XLSX.
        Document docxDocument = new Document(docxPath);
        string xlsxPath = "output.xlsx";
        docxDocument.Save(xlsxPath, SaveFormat.Xlsx);

        if (!File.Exists(xlsxPath) || new FileInfo(xlsxPath).Length == 0)
            throw new InvalidOperationException("XLSX file was not created successfully.");

        Console.WriteLine("Conversion completed: PDF -> DOCX -> XLSX.");
    }
}
