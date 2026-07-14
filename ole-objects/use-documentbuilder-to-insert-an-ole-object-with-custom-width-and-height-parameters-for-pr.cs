using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and attach a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a description before the OLE object.
        builder.Writeln("Embedded OLE object with custom width and height:");

        // Create a simple text file in memory to embed as an OLE package.
        byte[] fileContent = System.Text.Encoding.UTF8.GetBytes("Sample embedded text file content.");
        using (MemoryStream oleStream = new MemoryStream(fileContent))
        {
            // Insert the OLE object. "Package" is a generic progId for unknown file types.
            // The object is inserted as an icon (asIcon = true) with no custom presentation image.
            Shape oleShape = builder.InsertOleObject(oleStream, "Package", true, null);

            // Set precise layout dimensions (points). 1 point = 1/72 inch.
            oleShape.Width = 200;   // ~2.78 inches
            oleShape.Height = 100;  // ~1.39 inches
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "OleObjectCustomSize.docx");
        doc.Save(outputPath);
    }
}
