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

        // Prepare some dummy data to embed as an OLE object (a simple text file).
        byte[] oleData = System.Text.Encoding.UTF8.GetBytes("Hello, OLE object!");
        using (MemoryStream oleStream = new MemoryStream(oleData))
        {
            // Insert the OLE object into the document.
            // progId "Package" indicates a generic OLE package.
            // asIcon = false so the object is displayed as its content, not as an icon.
            Shape oleShape = builder.InsertOleObject(oleStream, "Package", false, null);

            // After insertion, retrieve the display width and height of the OLE shape (in points).
            double displayWidth = oleShape.Width;
            double displayHeight = oleShape.Height;

            // Store or use the dimensions for layout calculations.
            // For demonstration, write them to the console.
            Console.WriteLine($"OLE object display width: {displayWidth} points");
            Console.WriteLine($"OLE object display height: {displayHeight} points");
        }

        // Save the document to the file system.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "OleObject.docx");
        doc.Save(outputPath);
    }
}
