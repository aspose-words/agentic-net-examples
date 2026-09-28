using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Path to the input Word document.
        string inputPath = "input.docx";

        // Ensure the input file exists to avoid a FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
            return;
        }

        // Load the document.
        Document doc = new Document(inputPath);

        // Counter for naming extracted OLE files.
        int oleIndex = 0;

        // Iterate through all shapes in the document.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Check if the shape contains an OLE object.
            if (shape.OleFormat != null)
            {
                // Define the output file name.
                string outputFile = $"OleObject_{oleIndex}.bin";

                // Save the OLE object to a binary file.
                shape.OleFormat.Save(outputFile);
                Console.WriteLine($"Saved OLE object to {outputFile}");

                oleIndex++;
            }
        }

        if (oleIndex == 0)
        {
            Console.WriteLine("No OLE objects were found in the document.");
        }
    }
}
