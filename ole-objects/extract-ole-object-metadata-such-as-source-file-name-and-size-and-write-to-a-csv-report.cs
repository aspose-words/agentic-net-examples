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
        // Prepare a sample document with an embedded OLE object.
        string docPath = "SampleDocument.docx";
        string oleSourcePath = "SampleText.txt";
        File.WriteAllText(oleSourcePath, "This is a sample text file for OLE embedding.");

        // Create a new document and insert the OLE object.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        using (FileStream oleStream = new FileStream(oleSourcePath, FileMode.Open, FileAccess.Read))
        {
            // Insert the OLE object as a package (generic OLE container).
            // Rule: InsertOleObject(Stream, string, bool, Stream)
            Shape oleShape = builder.InsertOleObject(oleStream, "Package", false, null);
            // Optionally set display name for the package.
            oleShape.OleFormat.OlePackage.FileName = Path.GetFileName(oleSourcePath);
            oleShape.OleFormat.OlePackage.DisplayName = Path.GetFileName(oleSourcePath);
        }
        // Save the document (rule: Document.Save)
        doc.Save(docPath);

        // Load the document (rule: Document constructor)
        Document loadedDoc = new Document(docPath);

        // Prepare CSV output.
        string csvPath = "OleMetadataReport.csv";
        var sb = new StringBuilder();
        sb.AppendLine("ShapeIndex,ProgId,IsLink,SourceFullName,SuggestedFileName,SuggestedExtension,SizeInBytes");

        // Iterate over all shapes and collect OLE metadata.
        var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true).OfType<Shape>();
        int index = 0;
        foreach (var shape in shapes)
        {
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue;

            // Gather metadata.
            string progId = oleFormat.ProgId ?? string.Empty;
            bool isLink = oleFormat.IsLink;
            string sourceFullName = oleFormat.SourceFullName ?? string.Empty;
            string suggestedFileName = oleFormat.SuggestedFileName ?? string.Empty;
            string suggestedExtension = oleFormat.SuggestedExtension ?? string.Empty;

            // Size: use raw data length for embedded objects; for linked objects size may be unknown.
            long size = 0;
            try
            {
                byte[] rawData = oleFormat.GetRawData();
                if (rawData != null)
                    size = rawData.Length;
            }
            catch
            {
                // If GetRawData throws, keep size as 0.
            }

            sb.AppendLine($"{index},{EscapeCsv(progId)},{isLink},{EscapeCsv(sourceFullName)},{EscapeCsv(suggestedFileName)},{EscapeCsv(suggestedExtension)},{size}");
            index++;
        }

        // Write CSV file (rule: File.WriteAllText)
        File.WriteAllText(csvPath, sb.ToString());

        // Clean up temporary files (optional).
        // File.Delete(oleSourcePath);
        // File.Delete(docPath);
    }

    // Helper to escape CSV fields that may contain commas or quotes.
    private static string EscapeCsv(string field)
    {
        if (field.Contains(',') || field.Contains('\"') || field.Contains('\n') || field.Contains('\r'))
        {
            string escaped = field.Replace("\"", "\"\"");
            return $"\"{escaped}\"";
        }
        return field;
    }
}
