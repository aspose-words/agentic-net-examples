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
        // Prepare a temporary file to be embedded as an OLE object.
        string tempFolder = Path.Combine(Directory.GetCurrentDirectory(), "Temp");
        Directory.CreateDirectory(tempFolder);
        string sampleFilePath = Path.Combine(tempFolder, "sample.txt");
        File.WriteAllText(sampleFilePath, "This is a sample text file for OLE embedding.");

        // Create a new Word document and insert the OLE object.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert the file as an embedded OLE object (not a link, not an icon).
        builder.InsertOleObject(sampleFilePath, false, false, null);

        // Collect OLE metadata.
        var csvBuilder = new StringBuilder();
        csvBuilder.AppendLine("ShapeIndex,SourceFullName,SizeBytes");

        var shapes = doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>().ToArray();
        for (int i = 0; i < shapes.Length; i++)
        {
            Shape shape = shapes[i];
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue; // Not an OLE object.

            // SourceFullName may be empty for embedded objects.
            string sourceName = oleFormat.SourceFullName ?? string.Empty;

            // Get raw data size. For linked objects this may throw; handle gracefully.
            long size = 0;
            try
            {
                byte[] rawData = oleFormat.GetRawData();
                size = rawData?.LongLength ?? 0;
            }
            catch (InvalidOperationException)
            {
                // Linked objects cannot provide raw data.
                size = 0;
            }

            csvBuilder.AppendLine($"{i},\"{sourceName}\",{size}");
        }

        // Write CSV report to file.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "OleMetadataReport.csv");
        File.WriteAllText(reportPath, csvBuilder.ToString(), Encoding.UTF8);

        // Clean up temporary files.
        try { File.Delete(sampleFilePath); } catch { }
        try { Directory.Delete(tempFolder, true); } catch { }

        // Indicate completion (no interactive prompts).
        Console.WriteLine($"OLE metadata report generated at: {reportPath}");
    }
}
