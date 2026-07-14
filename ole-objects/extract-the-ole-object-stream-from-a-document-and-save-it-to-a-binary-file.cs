using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare a simple text file that will be embedded as an OLE package.
        string sampleFilePath = Path.Combine(Directory.GetCurrentDirectory(), "sample.txt");
        File.WriteAllText(sampleFilePath, "This is sample content for the OLE object.");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the OLE object (as a package) into the document.
        using (FileStream sampleFileStream = new FileStream(sampleFilePath, FileMode.Open, FileAccess.Read))
        {
            // progId "Package" indicates a generic OLE package.
            // asIcon = false to embed the content directly.
            // presentation = null uses the default presentation.
            builder.InsertOleObject(sampleFileStream, "Package", false, null);
        }

        // Retrieve the first shape that contains the OLE object.
        Shape oleShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
        OleFormat oleFormat = oleShape.OleFormat;

        // Define the output binary file path for the extracted OLE data.
        string outputOlePath = Path.Combine(Directory.GetCurrentDirectory(), "extracted_ole.bin");

        // Save the OLE object data to the binary file using a stream.
        using (FileStream outputStream = new FileStream(outputOlePath, FileMode.Create, FileAccess.Write))
        {
            oleFormat.Save(outputStream);
        }

        // Optional: indicate completion.
        Console.WriteLine("OLE object extracted to: " + outputOlePath);
    }
}
