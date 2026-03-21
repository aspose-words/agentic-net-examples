using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class RetrieveOleRawData
{
    static void Main()
    {
        // Resolve the document path relative to the current working directory.
        string docPath = Path.Combine(Environment.CurrentDirectory, "OleObjects.docx");

        // If the file does not exist, create an empty document so the program can run without error.
        if (!File.Exists(docPath))
        {
            Console.WriteLine($"Document not found at '{docPath}'. Creating an empty placeholder document.");
            var emptyDoc = new Document();
            emptyDoc.Save(docPath);
        }

        // Load the Word document that contains OLE objects.
        Document doc = new Document(docPath);

        // Iterate through all shapes in the document.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Check if the shape contains an OLE object.
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue; // Not an OLE shape, skip.

            // Determine whether the OLE object is embedded or linked.
            Console.WriteLine($"OLE object is {(oleFormat.IsLink ? "linked" : "embedded")}.");

            // Retrieve the raw binary data of the embedded OLE object.
            // For linked objects this method throws; handle accordingly.
            if (!oleFormat.IsLink)
            {
                byte[] rawData = oleFormat.GetRawData();

                // Example custom processing: display size and optionally save to a file.
                Console.WriteLine($"Raw data length: {rawData.Length} bytes.");

                // Save the raw data to a file in a safe output folder.
                string suggestedFileName = string.IsNullOrEmpty(oleFormat.SuggestedFileName)
                    ? "OleObject.bin"
                    : oleFormat.SuggestedFileName;

                string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
                Directory.CreateDirectory(outputDir);

                string outputPath = Path.Combine(outputDir, suggestedFileName);
                File.WriteAllBytes(outputPath, rawData);
                Console.WriteLine($"Raw data saved to: {outputPath}");
            }
            else
            {
                Console.WriteLine("Skipping linked OLE object; raw data not available.");
            }
        }

        Console.WriteLine("Processing completed.");
    }
}
