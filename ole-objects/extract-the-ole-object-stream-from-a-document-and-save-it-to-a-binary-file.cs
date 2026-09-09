using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Prepare some sample data to embed as an OLE object.
        byte[] sampleData = System.Text.Encoding.UTF8.GetBytes("Sample OLE object data");
        using (MemoryStream dataStream = new MemoryStream(sampleData))
        {
            // Insert the OLE object into the document.
            // progId "Package" indicates a generic OLE package.
            // asIcon = false (display content), presentation = null (default icon if needed).
            DocumentBuilder builder = new DocumentBuilder(doc);
            Shape oleShape = builder.InsertOleObject(dataStream, "Package", false, null);
            
            // Access the OleFormat of the inserted shape.
            OleFormat oleFormat = oleShape.OleFormat;

            // Define the output file path for the extracted OLE stream.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ExtractedOle.bin");

            // Save the OLE object data to a binary file using a stream.
            using (FileStream outputStream = new FileStream(outputPath, FileMode.Create))
            {
                oleFormat.Save(outputStream);
            }

            // Optional: indicate completion.
            Console.WriteLine($"OLE object extracted to: {outputPath}");
        }
    }
}
