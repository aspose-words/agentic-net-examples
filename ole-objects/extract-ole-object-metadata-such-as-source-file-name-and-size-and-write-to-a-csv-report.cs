using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Path to the Word document that may contain OLE objects.
        string docPath = "InputDocument.docx";

        // Path for the CSV report that will be generated.
        string csvPath = "OleMetadataReport.csv";

        // Load the document if it exists; otherwise create an empty document.
        Document doc = File.Exists(docPath) ? new Document(docPath) : new Document();

        // Prepare a StringBuilder to compose CSV content.
        // Header: SourceFileName,SizeInBytes
        StringBuilder csvBuilder = new StringBuilder();
        csvBuilder.AppendLine("SourceFileName,SizeInBytes");

        // Iterate through all shapes in the document.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Check if the shape contains an OLE object.
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue;

            // Retrieve the source file name. For embedded objects this may be empty.
            string sourceFileName = oleFormat.SourceFullName ?? string.Empty;

            // Retrieve the raw data of the OLE object to determine its size.
            // GetRawData returns a byte array; its Length is the size in bytes.
            byte[] rawData = oleFormat.GetRawData();
            long sizeInBytes = rawData?.LongLength ?? 0;

            // Escape any double quotes in the file name and wrap it in quotes.
            string escapedFileName = $"\"{sourceFileName.Replace("\"", "\"\"")}\"";

            // Append a line to the CSV.
            csvBuilder.AppendLine($"{escapedFileName},{sizeInBytes}");
        }

        // Write the CSV content to the output file.
        File.WriteAllText(csvPath, csvBuilder.ToString(), Encoding.UTF8);
    }
}
