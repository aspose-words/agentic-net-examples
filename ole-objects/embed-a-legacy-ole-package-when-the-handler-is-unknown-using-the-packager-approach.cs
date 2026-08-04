using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare some data to embed as an OLE package (e.g., a simple text file).
        byte[] fileBytes = System.Text.Encoding.UTF8.GetBytes("Hello from embedded OLE package!");
        using (MemoryStream stream = new MemoryStream(fileBytes))
        {
            // Insert the OLE object using the legacy "Package" progId.
            // The object is inserted as an icon (asIcon = true) with the default presentation.
            Shape shape = builder.InsertOleObject(stream, "Package", true, null);

            // Set the OLE package's file name and display name.
            shape.OleFormat.OlePackage.FileName = "Sample.txt";
            shape.OleFormat.OlePackage.DisplayName = "Sample Text File";
        }

        // Define output path and ensure the directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "OlePackageExample.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
