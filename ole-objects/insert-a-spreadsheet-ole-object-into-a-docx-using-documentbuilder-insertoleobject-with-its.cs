using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a temporary minimal Excel file (ZIP header) to act as the OLE data source.
        string tempExcelPath = Path.Combine(Path.GetTempPath(), "DummySpreadsheet.xlsx");
        File.WriteAllBytes(tempExcelPath, new byte[] { 0x50, 0x4B, 0x03, 0x04 }); // Minimal ZIP header.

        // Path where the resulting DOCX will be saved.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "SpreadsheetOleObject.docx");

        // Open the temporary Excel file as a stream and insert it as an OLE object.
        using (FileStream excelStream = File.OpenRead(tempExcelPath))
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a description paragraph.
            builder.Writeln("Spreadsheet OLE object:");

            // Insert the OLE object using its ProgId ("Excel.Sheet").
            // asIcon = false (display the content), presentation = null (use default icon if needed).
            builder.InsertOleObject(excelStream, "Excel.Sheet", false, null);

            // Save the document.
            doc.Save(outputPath);
        }

        // Clean up the temporary Excel file.
        if (File.Exists(tempExcelPath))
        {
            File.Delete(tempExcelPath);
        }
    }
}
