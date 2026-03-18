using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractOleObject
{
    static void Main()
    {
        // Path to the Word document that contains the OLE object.
        string documentPath = @"C:\Docs\Sample.docx";

        // Path where the extracted OLE binary stream will be saved.
        string outputPath = @"C:\Output\OleObject.bin";

        // Load the document from the file system.
        Document doc = new Document(documentPath);

        // Retrieve the first shape that holds an OLE object.
        Shape oleShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
        if (oleShape == null || oleShape.OleFormat == null)
        {
            Console.WriteLine("No OLE object found in the document.");
            return;
        }

        // Access the OleFormat of the shape.
        OleFormat oleFormat = oleShape.OleFormat;

        // Save the OLE object's data to a binary file using a stream.
        using (FileStream fileStream = new FileStream(outputPath, FileMode.Create))
        {
            oleFormat.Save(fileStream);
        }

        Console.WriteLine($"OLE object extracted successfully to: {outputPath}");
    }
}
