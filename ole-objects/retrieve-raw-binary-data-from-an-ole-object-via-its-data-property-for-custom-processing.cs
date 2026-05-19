using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Determine the base directory of the running assembly (e.g., bin/Debug/net8.0).
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;

        // The sample document is expected to be located in the project’s “Data” folder,
        // which is two levels up from the binary output folder.
        string projectRoot = Path.GetFullPath(Path.Combine(baseDir, "..", ".."));
        string inputPath = Path.Combine(projectRoot, "Data", "OLE objects.docx");

        // Verify that the input file exists; if not, inform the user and exit gracefully.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        // Load the Word document.
        Document doc = new Document(inputPath);

        // Prepare the output directory (project root + "Output").
        string outputDir = Path.Combine(projectRoot, "Output");
        Directory.CreateDirectory(outputDir);

        // Iterate through all shapes in the document and extract OLE objects.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>())
        {
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue; // Not an OLE object.

            // Retrieve the raw binary data of the OLE object.
            byte[] rawData = oleFormat.GetRawData();

            // Determine a suitable file extension (default to .bin if none is suggested).
            string extension = string.IsNullOrEmpty(oleFormat.SuggestedExtension)
                ? ".bin"
                : oleFormat.SuggestedExtension;

            // Build a unique file name for the extracted OLE data.
            string outputFileName = $"OleObject_{shape.GetHashCode()}{extension}";
            string outputPath = Path.Combine(outputDir, outputFileName);

            // Write the raw data to the file.
            File.WriteAllBytes(outputPath, rawData);
            Console.WriteLine($"Extracted OLE object saved to: {outputPath}");
        }
    }
}
