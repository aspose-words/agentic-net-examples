using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Input Word document that may contain OLE objects.
        const string inputPath = "Input.docx";

        // Output CSV file that will contain the OLE metadata.
        const string outputCsv = "OleMetadataReport.csv";

        // Load the document. If the file does not exist, create an empty document instead.
        Document doc;
        if (File.Exists(inputPath))
        {
            doc = new Document(inputPath);
        }
        else
        {
            // Create a new blank document to avoid a FileNotFoundException.
            doc = new Document();
        }

        // Prepare a writer for the CSV report.
        using (StreamWriter writer = new StreamWriter(outputCsv))
        {
            // Write CSV header.
            writer.WriteLine("ShapeIndex,SourceFileName,SizeBytes");

            // Retrieve all Shape nodes in the document.
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            // Iterate through each shape and extract OLE metadata when applicable.
            for (int i = 0; i < shapes.Count; i++)
            {
                Shape shape = (Shape)shapes[i];
                OleFormat ole = shape.OleFormat;

                // Skip shapes that do not contain OLE data.
                if (ole == null)
                    continue;

                // Determine the source file name.
                // For linked OLE objects this is SourceFullName; for embedded objects use SuggestedFileName.
                string sourceFileName = !string.IsNullOrEmpty(ole.SourceFullName)
                    ? ole.SourceFullName
                    : ole.SuggestedFileName ?? string.Empty;

                // Get the raw data of the OLE object to calculate its size.
                byte[] rawData = ole.GetRawData();
                long sizeBytes = rawData?.Length ?? 0;

                // Write a CSV line with the collected information.
                writer.WriteLine($"{i},\"{sourceFileName}\",{sizeBytes}");
            }
        }

        // Save the (potentially unchanged) document back to disk.
        doc.Save("Output.docx");
    }
}
