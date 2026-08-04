using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary minimal file to act as the spreadsheet data.
        string tempFilePath = Path.Combine(Directory.GetCurrentDirectory(), "temp.xlsx");
        // Write a few bytes that form the beginning of a ZIP archive (XLSX files are ZIP packages).
        File.WriteAllBytes(tempFilePath, new byte[] { 0x50, 0x4B, 0x03, 0x04 });

        // Path where the resulting DOCX will be saved.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OleObject.docx");

        // Open the temporary file as a stream and insert it as an OLE object.
        using (FileStream spreadsheetStream = File.OpenRead(tempFilePath))
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a description paragraph.
            builder.Writeln("Spreadsheet OLE object:");

            // Insert the OLE object using its ProgId. The object will be displayed as its content.
            // Parameters: stream, progId, asIcon (false = show content), presentation (null = default image).
            builder.InsertOleObject(spreadsheetStream, "Excel.Sheet", false, null);

            // Save the document.
            doc.Save(outputPath);
        }

        // Clean up the temporary file.
        if (File.Exists(tempFilePath))
        {
            File.Delete(tempFilePath);
        }

        // Indicate completion (optional, not required for non‑interactive execution).
        Console.WriteLine("Document created at: " + outputPath);
    }
}
