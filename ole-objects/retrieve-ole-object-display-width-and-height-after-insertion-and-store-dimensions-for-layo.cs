using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare OLE object data (a simple text file) in a memory stream.
        byte[] oleData = Encoding.UTF8.GetBytes("Hello, OLE object!");
        using (MemoryStream oleStream = new MemoryStream(oleData))
        {
            // Insert the OLE object into the document.
            // ProgId "Package" is used for generic OLE packages.
            // asIcon = false to display the content, presentation = null for default appearance.
            Shape oleShape = builder.InsertOleObject(oleStream, "Package", false, null);

            // Retrieve the display width and height of the inserted OLE object (in points).
            double displayWidth = oleShape.Width;
            double displayHeight = oleShape.Height;

            // Store dimensions for later layout calculations (example variables).
            double storedWidth = displayWidth;
            double storedHeight = displayHeight;

            // Output the dimensions to the console.
            Console.WriteLine($"OLE object display width: {storedWidth} points");
            Console.WriteLine($"OLE object display height: {storedHeight} points");
        }

        // Save the document to a file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OleObjectExample.docx");
        doc.Save(outputPath);
    }
}
