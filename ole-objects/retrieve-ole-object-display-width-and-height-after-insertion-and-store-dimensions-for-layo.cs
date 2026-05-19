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

        // Prepare a simple byte array to act as OLE package data.
        byte[] dummyData = new byte[] { 0x00 };
        using (MemoryStream oleStream = new MemoryStream(dummyData))
        {
            // Insert the OLE object into the document.
            // progId "Package" creates a generic OLE package.
            // asIcon = false (display as content), presentation = null (default icon if needed).
            Shape oleShape = builder.InsertOleObject(oleStream, "Package", false, null);
            
            // After insertion, the shape's Width and Height represent the display size in points.
            double displayWidth = oleShape.Width;   // Width in points (1 point = 1/72 inch)
            double displayHeight = oleShape.Height; // Height in points

            // Store dimensions for later layout calculations.
            // Example: calculate the area in square points.
            double area = displayWidth * displayHeight;

            // Output the retrieved dimensions.
            Console.WriteLine($"OLE object display width: {displayWidth} pt");
            Console.WriteLine($"OLE object display height: {displayHeight} pt");
            Console.WriteLine($"Calculated area: {area} pt²");
        }

        // Save the document to the file system.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "OleObjectDemo.docx");
        doc.Save(outputPath);
    }
}
