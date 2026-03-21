using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;

class OleMetadataExtractor
{
    static void Main()
    {
        // Use paths relative to the executable directory.
        string baseDir = AppContext.BaseDirectory;
        string inputPath = Path.Combine(baseDir, "InputDocument.docx");
        string csvPath = Path.Combine(baseDir, "OleMetadataReport.csv");

        // Ensure the input document exists. If not, create a minimal document.
        if (!File.Exists(inputPath))
        {
            var emptyDoc = new Document();
            emptyDoc.Save(inputPath);
        }

        // Load the Word document.
        Document doc = new Document(inputPath);

        // Prepare a StringBuilder to build CSV content.
        var csvBuilder = new StringBuilder();
        csvBuilder.AppendLine("ShapeIndex,SourceFileName,SizeBytes");

        // Iterate over all Shape nodes in the document.
        var shapes = doc.GetChildNodes(NodeType.Shape, true);
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
                : ole.SuggestedFileName ?? string.Empty;

            // Get the raw data of the OLE object to determine its size in bytes.
            byte[] rawData = ole.GetRawData();
            long sizeBytes = rawData?.LongLength ?? 0;

            // Append a CSV line for this OLE object.
            csvBuilder.AppendLine($"{shapeIndex},\"{sourceFileName}\",{sizeBytes}");

            shapeIndex++;
        }

        // Write the CSV content to a file.
        File.WriteAllText(csvPath, csvBuilder.ToString(), Encoding.UTF8);

        Console.WriteLine($"Report written to: {csvPath}");
    }
}
