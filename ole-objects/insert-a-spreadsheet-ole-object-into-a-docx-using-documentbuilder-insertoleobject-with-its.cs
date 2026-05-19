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

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare dummy spreadsheet data.
        byte[] dummyData = System.Text.Encoding.UTF8.GetBytes("Dummy spreadsheet content");

        // Insert the OLE object using a stream and the Excel ProgId.
        using (MemoryStream stream = new MemoryStream(dummyData))
        {
            // asIcon = false (display content), presentation = null (default icon if needed).
            builder.InsertOleObject(stream, "Excel.Sheet", false, null);
        }

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "SpreadsheetOleObject.docx");
        doc.Save(outputPath);
    }
}
