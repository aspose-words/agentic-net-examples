using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary file to embed as an OLE object.
        string tempFolder = Path.GetTempPath();
        string tempFilePath = Path.Combine(tempFolder, "SampleText.txt");
        File.WriteAllText(tempFilePath, "This is a sample text file for OLE embedding.");

        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the OLE object from the temporary file as a package (generic OLE object).
        // The InsertOleObject method returns the Shape that contains the OLE object.
        Shape oleShape;
        using (FileStream fs = new FileStream(tempFilePath, FileMode.Open, FileAccess.Read))
        {
            oleShape = builder.InsertOleObject(fs, "Package", false, null);
        }

        // Retrieve the display width and height of the OLE object (in points).
        double displayWidth = oleShape.Width;
        double displayHeight = oleShape.Height;

        // Store dimensions for further layout calculations (example: calculate area).
        double oleArea = displayWidth * displayHeight;

        // Output the dimensions to the console (no user interaction required).
        Console.WriteLine($"OLE object width: {displayWidth} pt");
        Console.WriteLine($"OLE object height: {displayHeight} pt");
        Console.WriteLine($"Calculated area: {oleArea} pt²");

        // Save the document to a temporary location.
        string outputPath = Path.Combine(tempFolder, "OleObjectDocument.docx");
        doc.Save(outputPath);
    }
}
