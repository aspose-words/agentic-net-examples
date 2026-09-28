using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a simple table.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Build the table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Row1Col1");
        builder.InsertCell();
        builder.Write("Row1Col2");
        builder.EndRow();

        builder.EndTable();

        // Save the document as PDF.
        string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // Load the PDF and convert it to DOCX.
        Document pdfDoc = new Document(pdfPath);
        string docxPath = "sample.docx";
        pdfDoc.Save(docxPath, SaveFormat.Docx);
        if (!File.Exists(docxPath))
            throw new InvalidOperationException("DOCX file was not created.");

        // Load the DOCX and convert it to XLSX (tables become worksheets).
        Document docxDoc = new Document(docxPath);
        string xlsxPath = "sample.xlsx";
        docxDoc.Save(xlsxPath, SaveFormat.Xlsx);
        if (!File.Exists(xlsxPath))
            throw new InvalidOperationException("XLSX file was not created.");

        // Verify that the XLSX file contains data.
        FileInfo xlsxInfo = new FileInfo(xlsxPath);
        if (xlsxInfo.Length == 0)
            throw new InvalidOperationException("XLSX file is empty.");
    }
}
