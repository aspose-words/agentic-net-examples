using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main(string[] args)
    {
        // Paths for input document and output CSV report.
        const string inputPath = "input.docx";
        const string csvPath = "OleReport.csv";

        // Ensure the input document exists; create an empty one if it does not.
        if (!File.Exists(inputPath))
        {
            Document emptyDoc = new Document();
            emptyDoc.Save(inputPath);
        }

        // Load the Word document.
        Document doc = new Document(inputPath);

        // Retrieve all Shape nodes that contain an OLE object (OleFormat is not null).
        var oleShapes = doc.GetChildNodes(NodeType.Shape, true)
                           .Cast<Shape>()
                           .Where(s => s.OleFormat != null)
                           .ToList();

        // Prepare CSV content.
        StringBuilder csvBuilder = new StringBuilder();
        csvBuilder.AppendLine("Index,OriginalFileName,SizeBytes");

        for (int i = 0; i < oleShapes.Count; i++)
        {
            Shape shape = oleShapes[i];
            OleFormat oleFormat = shape.OleFormat;

            // Attempt to obtain the original file name via reflection (property may be absent in some versions).
            string originalFileName = string.Empty;
            PropertyInfo fileNameProp = oleFormat.GetType().GetProperty("OriginalFileName", BindingFlags.Public | BindingFlags.Instance);
            if (fileNameProp != null)
            {
                object value = fileNameProp.GetValue(oleFormat);
                originalFileName = value as string ?? string.Empty;
            }

            // Retrieve the OLE data bytes using the GetOleData method (available in all supported versions).
            byte[] oleData = null;
            MethodInfo getDataMethod = oleFormat.GetType().GetMethod("GetOleData", BindingFlags.Public | BindingFlags.Instance);
            if (getDataMethod != null)
            {
                object result = getDataMethod.Invoke(oleFormat, null);
                oleData = result as byte[];
            }

            long sizeBytes = oleData?.Length ?? 0;

            // Append a line to the CSV, quoting the file name to handle commas.
            csvBuilder.AppendLine($"{i + 1},\"{originalFileName}\",{sizeBytes}");
        }

        // Write the CSV report to disk.
        File.WriteAllText(csvPath, csvBuilder.ToString());

        // Confirmation message.
        Console.WriteLine($"OLE metadata extraction complete. Report saved to '{csvPath}'.");
    }
}
