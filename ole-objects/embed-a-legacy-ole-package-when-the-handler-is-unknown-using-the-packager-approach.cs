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

        // Prepare some sample data to embed as an OLE package.
        byte[] sampleData = System.Text.Encoding.UTF8.GetBytes("This is the content of the embedded file.");
        using (MemoryStream stream = new MemoryStream(sampleData))
        {
            // Insert the OLE object using the legacy "Package" progId.
            // The object will be displayed as an icon.
            Shape oleShape = builder.InsertOleObject(stream, "Package", true, null);

            // Configure the OLE package properties.
            oleShape.OleFormat.OlePackage.FileName = "SampleFile.txt";
            oleShape.OleFormat.OlePackage.DisplayName = "Sample File.txt";
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);
    }
}
