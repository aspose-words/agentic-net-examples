using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // 1. Create a sample Word document containing a table.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Build a simple 2‑column table with a header row and two data rows.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        for (int i = 1; i <= 2; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i} Col 1");
            builder.InsertCell();
            builder.Write($"Row {i} Col 2");
            builder.EndRow();
        }
        builder.EndTable();

        // 2. Save the document as PDF – this will be the source PDF for conversion.
        string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);
        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
            throw new InvalidOperationException("PDF file was not created.");

        // 3. Load the PDF and convert it to DOCX.
        Document pdfDoc = new Document(pdfPath);
        string docxPath = "converted.docx";
        pdfDoc.Save(docxPath, SaveFormat.Docx);
        if (!File.Exists(docxPath) || new FileInfo(docxPath).Length == 0)
            throw new InvalidOperationException("DOCX file was not created.");

        // 4. Load the DOCX and convert it to XLSX, extracting tables into a spreadsheet.
        Document docxDoc = new Document(docxPath);
        string xlsxPath = "tables.xlsx";

        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            SaveFormat = SaveFormat.Xlsx,
            // Optional: place all sections on a single worksheet.
            SectionMode = XlsxSectionMode.SingleWorksheet
        };

        docxDoc.Save(xlsxPath, xlsxOptions);
        if (!File.Exists(xlsxPath) || new FileInfo(xlsxPath).Length == 0)
            throw new InvalidOperationException("XLSX file was not created.");

        // Conversion sequence completed successfully.
    }
}
