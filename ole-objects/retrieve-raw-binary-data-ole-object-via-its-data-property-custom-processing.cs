using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class RetrieveOleRawData
{
    static void Main()
    {
        // Load the Word document that contains OLE objects.
        // Replace with the actual path to your document.
        string docPath = @"C:\Docs\OleObjects.docx";
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

                // Optional: save the raw data to a file using the suggested filename.
                string suggestedFileName = oleFormat.SuggestedFileName ?? "OleObject.bin";
                string outputPath = Path.Combine(@"C:\Output", suggestedFileName);
                File.WriteAllBytes(outputPath, rawData);
                Console.WriteLine($"Raw data saved to: {outputPath}");
            }
            else
            {
                Console.WriteLine("Skipping linked OLE object; raw data not available.");
            }
        }
    }
}
