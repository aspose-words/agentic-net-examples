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

        // Prepare a temporary file to be used as the OLE object source.
        string tempFilePath = Path.Combine(Path.GetTempPath(), "sample.txt");
        File.WriteAllText(tempFilePath, "This is a sample OLE object content.");

        // Insert the OLE object into the document.
        // The overload requires a fourth parameter for the icon file name; we pass null to use the default icon.
        Shape oleShape = builder.InsertOleObject(tempFilePath, false, false, null);

        // Retrieve the current display width and height of the OLE object (in points).
        double originalWidth = oleShape.Width;
        double originalHeight = oleShape.Height;

        // Output the original dimensions to the console.
        Console.WriteLine($"Original OLE size: Width = {originalWidth} pt, Height = {originalHeight} pt");

        // Adjust the size of the OLE object (e.g., double its dimensions).
        oleShape.Width = originalWidth * 2;
        oleShape.Height = originalHeight * 2;

        // Output the new dimensions to the console.
        Console.WriteLine($"Adjusted OLE size: Width = {oleShape.Width} pt, Height = {oleShape.Height} pt");

        // Save the document to a file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");

        // Clean up the temporary file.
        if (File.Exists(tempFilePath))
        {
            File.Delete(tempFilePath);
        }
    }
}
