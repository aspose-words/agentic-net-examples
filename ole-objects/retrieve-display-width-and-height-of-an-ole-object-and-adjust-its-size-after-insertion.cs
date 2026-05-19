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

        // Prepare a temporary text file to embed as an OLE object.
        string tempFilePath = Path.Combine(Path.GetTempPath(), "SampleText.txt");
        File.WriteAllText(tempFilePath, "This is a sample text file for OLE embedding.");

        // Insert the OLE object (embedded, not as an icon).
        Shape oleShape = builder.InsertOleObject(tempFilePath, false, false, null);

        // Retrieve the current display width and height (in points).
        double originalWidth = oleShape.Width;
        double originalHeight = oleShape.Height;

        Console.WriteLine($"Original OLE display size: Width = {originalWidth} pt, Height = {originalHeight} pt");

        // Adjust the size of the OLE object – for example, double its dimensions.
        oleShape.Width = originalWidth * 2;
        oleShape.Height = originalHeight * 2;

        Console.WriteLine($"Adjusted OLE display size: Width = {oleShape.Width} pt, Height = {oleShape.Height} pt");

        // Save the document to the output file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OleObjectAdjusted.docx");
        doc.Save(outputPath);
    }
}
