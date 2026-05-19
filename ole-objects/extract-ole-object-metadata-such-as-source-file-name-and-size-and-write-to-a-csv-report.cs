using System;
using System.IO;
using System.Text;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Input and output file paths.
        string inputPath = "InputDocument.docx";
        string csvReportPath = "OleMetadataReport.csv";

        // If the input document does not exist, create a minimal document so the program can run.
        if (!File.Exists(inputPath))
        {
            // Create a new blank document and add a simple paragraph.
            Document emptyDoc = new Document();
            emptyDoc.FirstSection.Body.FirstParagraph.AppendChild(new Run(emptyDoc, "This document contains no OLE objects."));
            emptyDoc.Save(inputPath);
        }

        // Load the Word document.
        Document doc = new Document(inputPath);

        // Prepare CSV content.
        StringBuilder csvBuilder = new StringBuilder();
        csvBuilder.AppendLine("Index,SourceFileName,SizeBytes");

        // Find all shapes that contain OLE objects.
        var oleShapes = doc.GetChildNodes(NodeType.Shape, true)
                           .OfType<Shape>()
                           .Where(s => s.OleFormat != null)
                           .ToList();

        for (int i = 0; i < oleShapes.Count; i++)
        {
            Shape shape = oleShapes[i];
            OleFormat ole = shape.OleFormat;

            // Determine the source file name.
            string sourceFileName = string.IsNullOrEmpty(ole.SourceFullName)
                ? ole.SuggestedFileName ?? "EmbeddedObject"
                : Path.GetFileName(ole.SourceFullName);

            // Get the size of the raw OLE data.
            long sizeBytes = ole.GetRawData()?.Length ?? 0;

            // Append a line to the CSV.
            csvBuilder.AppendLine($"{i + 1},\"{sourceFileName}\",{sizeBytes}");
        }

        // Write the CSV report to disk.
        File.WriteAllText(csvReportPath, csvBuilder.ToString(), Encoding.UTF8);
    }
}
