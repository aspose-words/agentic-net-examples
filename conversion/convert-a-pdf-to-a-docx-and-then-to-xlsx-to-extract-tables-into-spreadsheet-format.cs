using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample Word document with a table.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Build a simple 2x3 table.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // First data row.
        builder.InsertCell();
        builder.Write("Row1 Col1");
        builder.InsertCell();
        builder.Write("Row1 Col2");
        builder.EndRow();

        // Second data row.
        builder.InsertCell();
        builder.Write("Row2 Col1");
        builder.InsertCell();
        builder.Write("Row2 Col2");
        builder.EndRow();

        builder.EndTable();

        // Save the document as PDF (this will be our input PDF).
        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // Step 2: Load the PDF and convert it to DOCX.
        Document pdfDoc = new Document(pdfPath);
        const string docxPath = "converted.docx";
        pdfDoc.Save(docxPath, SaveFormat.Docx);
        if (!File.Exists(docxPath))
            throw new InvalidOperationException("DOCX file was not created.");

        // Step 3: Load the DOCX and convert it to XLSX (tables become worksheets).
        Document docxDoc = new Document(docxPath);
        const string xlsxPath = "tables.xlsx";

        // Use default XlsxSaveOptions; can be omitted, but shown for clarity.
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            SaveFormat = SaveFormat.Xlsx,
            SectionMode = XlsxSectionMode.MultipleWorksheets
        };

        docxDoc.Save(xlsxPath, xlsxOptions);
        if (!File.Exists(xlsxPath))
            throw new InvalidOperationException("XLSX file was not created.");

        // All conversions completed successfully.
        Console.WriteLine("Conversion pipeline completed. Files created:");
        Console.WriteLine(pdfPath);
        Console.WriteLine(docxPath);
        Console.WriteLine(xlsxPath);
    }
}
