using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;

class OleMetadataExtractor
{
    static void Main()
    {
        // Path to the input Word document.
        string inputPath = @"C:\Docs\InputDocument.docx";

        // Path to the output CSV report.
        string csvPath = @"C:\Docs\OleMetadataReport.csv";

        // Load the Word document (uses the provided load rule).
        Document doc = new Document(inputPath);

        // Prepare a StringBuilder to build CSV content.
        StringBuilder csvBuilder = new StringBuilder();
        csvBuilder.AppendLine("ShapeIndex,SourceFileName,SizeBytes");

        // Iterate over all Shape nodes in the document.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int shapeIndex = 0;

        foreach (Shape shape in shapes)
        {
            // Only process shapes that contain an OLE object.
            OleFormat ole = shape.OleFormat;
            if (ole == null)
                continue;

            // SourceFullName may be empty for embedded objects; use SuggestedFileName instead.
            string sourceFileName = !string.IsNullOrEmpty(ole.SourceFullName)
                ? Path.GetFileName(ole.SourceFullName)
                : ole.SuggestedFileName;

            // Get the raw data of the OLE object to determine its size in bytes.
            byte[] rawData = ole.GetRawData();
            long sizeBytes = rawData?.LongLength ?? 0;

            // Append a CSV line for this OLE object.
            csvBuilder.AppendLine($"{shapeIndex},\"{sourceFileName}\",{sizeBytes}");

            shapeIndex++;
        }

        // Write the CSV content to a file (uses the provided save rule).
        File.WriteAllText(csvPath, csvBuilder.ToString(), Encoding.UTF8);
    }
}
