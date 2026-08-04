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

        // Prepare some dummy data to embed as an OLE package.
        byte[] dummyData = System.Text.Encoding.UTF8.GetBytes("Hello, Aspose.Words OLE object!");
        using (MemoryStream oleStream = new MemoryStream(dummyData))
        {
            // Insert the OLE object into the document.
            // Parameters: stream, progId ("Package" for generic OLE package), asIcon = false, presentation = null.
            Shape oleShape = builder.InsertOleObject(oleStream, "Package", false, null);

            // Retrieve the current display size of the OLE object (in points).
            double originalWidth = oleShape.Width;
            double originalHeight = oleShape.Height;

            Console.WriteLine($"Original OLE size: Width = {originalWidth} pt, Height = {originalHeight} pt");

            // Adjust the size of the OLE object – for example, increase both dimensions by 50%.
            oleShape.Width = originalWidth * 1.5;
            oleShape.Height = originalHeight * 1.5;

            Console.WriteLine($"Adjusted OLE size: Width = {oleShape.Width} pt, Height = {oleShape.Height} pt");
        }

        // Save the document to the file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OleObjectAdjusted.docx");
        doc.Save(outputPath);
    }
}
