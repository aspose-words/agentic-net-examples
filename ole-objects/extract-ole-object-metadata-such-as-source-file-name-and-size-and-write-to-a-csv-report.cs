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
        // Resolve the document path relative to the executable folder.
        string documentPath = Path.Combine(AppContext.BaseDirectory, "OleObjects.docx");

        // If the source document does not exist, create an empty one so the program can run without error.
        if (!File.Exists(documentPath))
        {
            // Create a blank document and save it to the expected location.
            Document emptyDoc = new Document();
            emptyDoc.Save(documentPath);
        }

        // Load the Word document (existing or newly created).
        Document doc = new Document(documentPath);

        // Prepare CSV content.
        StringBuilder csvBuilder = new StringBuilder();
        csvBuilder.AppendLine("SourceFileName,SizeInBytes");

        // Iterate over all shapes that may contain OLE objects.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>())
        {
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue; // Not an OLE object.

            // Determine a display name for the OLE object.
            string sourceFileName = string.IsNullOrEmpty(oleFormat.SourceFullName)
                ? oleFormat.SuggestedFileName ?? "EmbeddedObject"
                : Path.GetFileName(oleFormat.SourceFullName);

            // Get the size of the raw OLE data (0 for linked objects that cannot be read).
            long sizeInBytes = 0;
            try
            {
                byte[] rawData = oleFormat.GetRawData();
                sizeInBytes = rawData?.LongLength ?? 0;
            }
            catch
            {
                // Linked objects may throw; treat size as 0.
                sizeInBytes = 0;
            }

            csvBuilder.AppendLine($"{EscapeCsv(sourceFileName)},{sizeInBytes}");
        }

        // Write the CSV report.
        string csvPath = Path.Combine(AppContext.BaseDirectory, "OleMetadataReport.csv");
        File.WriteAllText(csvPath, csvBuilder.ToString());

        Console.WriteLine($"OLE metadata report saved to '{csvPath}'.");
    }

    // Escapes CSV fields that contain commas, quotes or newlines.
    private static string EscapeCsv(string field)
    {
        if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
        {
            string escaped = field.Replace("\"", "\"\"");
            return $"\"{escaped}\"";
        }
        return field;
    }
}
