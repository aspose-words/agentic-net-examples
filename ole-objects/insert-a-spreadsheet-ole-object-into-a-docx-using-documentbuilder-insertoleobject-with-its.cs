using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare a dummy spreadsheet file (Excel expects a .xlsx file).
        string tempSpreadsheetPath = Path.Combine(Path.GetTempPath(), "DummySpreadsheet.xlsx");
        // Write minimal content; the file does not need to be a valid Excel workbook for this demo.
        File.WriteAllText(tempSpreadsheetPath, "Dummy spreadsheet content");

        // Insert the OLE object using a stream and the Excel progId.
        using (FileStream spreadsheetStream = File.OpenRead(tempSpreadsheetPath))
        {
            builder.Writeln("Spreadsheet OLE object:");
            // progId "Excel.Sheet" identifies the object as an Excel spreadsheet.
            // asIcon = false (display the content), presentation = null (use default appearance).
            builder.InsertOleObject(spreadsheetStream, "Excel.Sheet", false, null);
        }

        // Save the resulting document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OleSpreadsheet.docx");
        doc.Save(outputPath);
    }
}
