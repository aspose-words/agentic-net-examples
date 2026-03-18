using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleSpreadsheet
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Path to the Excel file that will be embedded.
        string excelPath = @"C:\Temp\Sample.xlsx";

        // Open the Excel file as a stream.
        using (FileStream excelStream = new FileStream(excelPath, FileMode.Open, FileAccess.Read))
        {
            // Insert the OLE object using its ProgId.
            // "Excel.Sheet.12" is the ProgId for Excel 2007+ (.xlsx) files.
            // asIcon = false displays the spreadsheet content directly.
            // presentation = null lets Aspose.Words use the default appearance.
            Shape oleShape = builder.InsertOleObject(excelStream, "Excel.Sheet.12", false, null);
        }

        // Save the resulting document.
        doc.Save(@"C:\Temp\OleSpreadsheet.docx");
    }
}
