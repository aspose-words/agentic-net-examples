using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // 1. Create a sample Word document containing a table and save it as PDF.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Build a simple 2x2 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Value 1");
        builder.InsertCell();
        builder.Write("Value 2");
        builder.EndTable();

        string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // 2. Load the PDF and convert it to DOCX.
        Document pdfDoc = new Document(pdfPath);
        string docxPath = "sample.docx";
        pdfDoc.Save(docxPath, SaveFormat.Docx);
        if (!File.Exists(docxPath))
            throw new InvalidOperationException("DOCX file was not created.");

        // 3. Load the DOCX and convert it to XLSX (spreadsheet) to extract tables.
        Document docxDoc = new Document(docxPath);
        string xlsxPath = "sample.xlsx";

        // Use XlsxSaveOptions to specify XLSX format and worksheet handling.
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            SaveFormat = SaveFormat.Xlsx,
            SectionMode = XlsxSectionMode.SingleWorksheet
        };
        docxDoc.Save(xlsxPath, xlsxOptions);
        if (!File.Exists(xlsxPath))
            throw new InvalidOperationException("XLSX file was not created.");

        // Conversion sequence completed successfully.
    }
}
