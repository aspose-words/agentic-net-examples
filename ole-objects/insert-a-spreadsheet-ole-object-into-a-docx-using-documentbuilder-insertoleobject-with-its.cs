using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare a dummy Excel spreadsheet stream.
        // In a real scenario, this would be the content of an actual .xlsx file.
        byte[] dummyExcelData = new byte[] { 0x50, 0x4B, 0x03, 0x04 }; // ZIP header (XLSX files are ZIP archives)
        using (MemoryStream excelStream = new MemoryStream(dummyExcelData))
        {
            // Insert the OLE object using its ProgId.
            // ProgId "Excel.Sheet" tells Word the object is an Excel spreadsheet.
            // asIcon = false (display the content), presentation = null (use default icon if needed).
            builder.InsertOleObject(excelStream, "Excel.Sheet", false, null);
        }

        // Save the document to the file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "SpreadsheetOleObject.docx");
        doc.Save(outputPath);
    }
}
