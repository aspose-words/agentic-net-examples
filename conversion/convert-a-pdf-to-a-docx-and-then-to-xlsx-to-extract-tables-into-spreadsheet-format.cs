using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing a simple table.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Row 1, Col 1");
        builder.InsertCell();
        builder.Write("Row 1, Col 2");
        builder.EndTable();

        // Save the document as PDF.
        string pdfPath = "sample.pdf";
        source.Save(pdfPath, SaveFormat.Pdf);
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // Load the PDF and convert it to DOCX.
        Document pdfDoc = new Document(pdfPath);
        string docxPath = "sample.docx";
        pdfDoc.Save(docxPath, SaveFormat.Docx);
        if (!File.Exists(docxPath))
            throw new InvalidOperationException("DOCX file was not created.");

        // Load the DOCX and convert it to XLSX (tables become worksheets).
        Document docx = new Document(docxPath);
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            SaveFormat = SaveFormat.Xlsx,
            SectionMode = XlsxSectionMode.MultipleWorksheets
        };
        string xlsxPath = "tables.xlsx";
        docx.Save(xlsxPath, xlsxOptions);
        if (!File.Exists(xlsxPath))
            throw new InvalidOperationException("XLSX file was not created.");
    }
}
