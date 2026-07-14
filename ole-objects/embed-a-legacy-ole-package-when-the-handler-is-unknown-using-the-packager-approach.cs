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

        // Prepare some sample data to embed as an OLE package.
        string sampleText = "Hello, this is embedded data.";
        byte[] sampleBytes = System.Text.Encoding.UTF8.GetBytes(sampleText);

        // Insert the OLE object using the legacy Packager approach (ProgID "Package").
        using (MemoryStream stream = new MemoryStream(sampleBytes))
        {
            // Insert as an icon (asIcon = true) with the default presentation.
            Shape oleShape = builder.InsertOleObject(stream, "Package", true, null);

            // Configure the OLE package properties.
            oleShape.OleFormat.OlePackage.FileName = "sample.txt";
            oleShape.OleFormat.OlePackage.DisplayName = "Sample Text File";
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);
    }
}
