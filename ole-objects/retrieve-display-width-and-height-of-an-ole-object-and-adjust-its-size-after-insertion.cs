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
        // Parameters: file name, isLinked = false, asIcon = false, presentation = null.
        Shape oleShape = builder.InsertOleObject(tempFilePath, false, false, null);

        // Retrieve the display width and height of the OLE object (in points).
        double originalWidth = oleShape.Width;
        double originalHeight = oleShape.Height;

        // Output the original dimensions.
        Console.WriteLine($"Original OLE display size: Width = {originalWidth} pt, Height = {originalHeight} pt");

        // Adjust the size of the OLE object – for example, increase both dimensions by 50%.
        oleShape.Width = originalWidth * 1.5;
        oleShape.Height = originalHeight * 1.5;

        // Output the new dimensions.
        Console.WriteLine($"Adjusted OLE display size: Width = {oleShape.Width} pt, Height = {oleShape.Height} pt");

        // Save the document to a temporary location.
        string outputPath = Path.Combine(Path.GetTempPath(), "OleObjectDemo.docx");
        doc.Save(outputPath);

        // Clean up the temporary text file.
        File.Delete(tempFilePath);
    }
}
