using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;   // Needed for the Table class

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample PDF containing a simple 2x2 table.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Build a 2x2 table.
        Table table = builder.StartTable();   // StartTable returns a Table object
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

        // Save the document as PDF – this will be the input file for conversion.
        string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 2. Load the PDF document.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // -----------------------------------------------------------------
        // 3. Configure XLSX save options to create a separate worksheet for each section.
        // -----------------------------------------------------------------
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            SectionMode = XlsxSectionMode.MultipleWorksheets
        };

        // -----------------------------------------------------------------
        // 4. Convert the PDF to XLSX.
        // -----------------------------------------------------------------
        string xlsxPath = "output.xlsx";
        pdfDoc.Save(xlsxPath, xlsxOptions);

        // -----------------------------------------------------------------
        // 5. Validate that the XLSX file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(xlsxPath))
        {
            throw new InvalidOperationException("Expected XLSX output was not created.");
        }

        Console.WriteLine("Conversion succeeded. XLSX file created at: " + Path.GetFullPath(xlsxPath));
    }
}
