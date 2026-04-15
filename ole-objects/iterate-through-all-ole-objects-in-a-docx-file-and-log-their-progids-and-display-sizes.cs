using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Path to the input DOCX file containing OLE objects.
        string inputPath = "input.docx";

        // Load the document.
        Document doc = new Document(inputPath);

        // Iterate through all shapes in the document.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Check if the shape contains an OLE object.
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat != null)
            {
                // Log the ProgId and the display size of the OLE object.
                Console.WriteLine($"OLE Object ProgId: {oleFormat.ProgId}, Size: {shape.Width}x{shape.Height} points");
            }
        }
    }
}
