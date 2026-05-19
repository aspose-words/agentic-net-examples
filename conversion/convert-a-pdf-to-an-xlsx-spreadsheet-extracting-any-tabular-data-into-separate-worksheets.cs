using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample Word document with a table.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Insert first table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        builder.EndTable();

        // Insert a page break to separate sections (each section will become a worksheet).
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Insert second table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Col A");
        builder.InsertCell();
        builder.Write("Col B");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Data A1");
        builder.InsertCell();
        builder.Write("Data B1");
        builder.EndRow();

        builder.EndTable();

        // Save the document as a PDF – this will be the input for conversion.
        const string pdfPath = "input.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Configure XLSX save options to create a separate worksheet for each section.
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            SectionMode = XlsxSectionMode.MultipleWorksheets,
            SaveFormat = SaveFormat.Xlsx
        };

        // Save the PDF content as an XLSX spreadsheet.
        const string xlsxPath = "output.xlsx";
        pdfDoc.Save(xlsxPath, xlsxOptions);

        // Validate that the XLSX file was created.
        if (!File.Exists(xlsxPath))
            throw new InvalidOperationException("Expected output XLSX was not created.");

        // Optional: clean up intermediate PDF if not needed.
        // File.Delete(pdfPath);
    }
}
