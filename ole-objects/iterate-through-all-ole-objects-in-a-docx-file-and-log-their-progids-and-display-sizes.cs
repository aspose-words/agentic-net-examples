using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Path to the input DOCX file.
        string inputPath = "Sample.docx";

        // Ensure the file exists. If it does not, create an empty document and save it.
        if (!File.Exists(inputPath))
        {
            Document emptyDoc = new Document(); // Create a blank document.
            emptyDoc.Save(inputPath);           // Save it so that the file exists for loading.
        }

        // Load the document.
        Document doc = new Document(inputPath);

        // Iterate through all shapes in the document.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>())
        {
            // Check if the shape contains an OLE object.
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat != null)
            {
                // Log the ProgId and display size (width and height in points).
                Console.WriteLine($"OLE Object ProgId: {oleFormat.ProgId}, Size: {shape.Width}x{shape.Height} points");
            }
        }
    }
}
