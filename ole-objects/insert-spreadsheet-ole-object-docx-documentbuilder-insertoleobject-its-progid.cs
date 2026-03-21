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

        // Create a minimal temporary Excel file (empty .xlsx) to embed.
        string excelPath = Path.Combine(Path.GetTempPath(), "Sample.xlsx");
        // Write a minimal ZIP file header so the file is a valid (though empty) .xlsx package.
        using (FileStream fs = new FileStream(excelPath, FileMode.Create, FileAccess.Write))
        {
            // End of central directory record for an empty zip archive.
            byte[] emptyZip = new byte[] { 0x50, 0x4B, 0x05, 0x06, 0x00, 0x00, 0x00, 0x00,
                                          0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                                          0x00, 0x00, 0x00, 0x00 };
            fs.Write(emptyZip, 0, emptyZip.Length);
        }

        // Open the Excel file as a stream.
        using (FileStream excelStream = new FileStream(excelPath, FileMode.Open, FileAccess.Read))
        {
            // Insert the OLE object using its ProgId.
            // "Excel.Sheet.12" is the ProgId for Excel 2007+ (.xlsx) files.
            // asIcon = false displays the spreadsheet content directly.
            // presentation = null lets Aspose.Words use the default appearance.
            Shape oleShape = builder.InsertOleObject(excelStream, "Excel.Sheet.12", false, null);
        }

        // Save the resulting document to the temporary folder.
        string outputPath = Path.Combine(Path.GetTempPath(), "OleSpreadsheet.docx");
        doc.Save(outputPath);

        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
