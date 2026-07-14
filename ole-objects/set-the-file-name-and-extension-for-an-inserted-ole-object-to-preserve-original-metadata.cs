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

        // Prepare a simple byte array that mimics a ZIP file (PK header).
        byte[] zipFileBytes = new byte[] { 0x50, 0x4B, 0x03, 0x04 };

        // Insert the OLE package from the byte array.
        using (MemoryStream stream = new MemoryStream(zipFileBytes))
        {
            // Insert as an OLE object displayed as an icon.
            Shape shape = builder.InsertOleObject(stream, "Package", true, null);

            // Set the original file name and display name to preserve metadata.
            shape.OleFormat.OlePackage.FileName = "Package file name.zip";
            shape.OleFormat.OlePackage.DisplayName = "Package display name.zip";
        }

        // Ensure the output directory exists.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Result.docx");
        doc.Save(outputPath);
    }
}
