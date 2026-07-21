using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare a simple file that will be embedded as an OLE package.
        string sourceFilePath = Path.Combine(Path.GetTempPath(), "SampleData.txt");
        File.WriteAllText(sourceFilePath, "This is sample data for OLE embedding.");

        // Load the file bytes into a memory stream.
        byte[] sourceBytes = File.ReadAllBytes(sourceFilePath);
        using (MemoryStream sourceStream = new MemoryStream(sourceBytes))
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert the OLE object (as a package) into the document.
            // progId "Package" indicates a generic OLE package.
            // asIcon = true makes it appear as an icon.
            Shape oleShape = builder.InsertOleObject(sourceStream, "Package", true, null);

            // Optionally set a display name for the package.
            oleShape.OleFormat.OlePackage.FileName = "SampleData.txt";
            oleShape.OleFormat.OlePackage.DisplayName = "SampleData.txt";

            // Save the document to a temporary location (not strictly required for extraction).
            string docPath = Path.Combine(Path.GetTempPath(), "OleDocument.docx");
            doc.Save(docPath);
        }

        // Load the document we just created.
        Document loadedDoc = new Document(sourceFilePath.Replace("SampleData.txt", "OleDocument.docx"));

        // Iterate through all shapes to find OLE objects.
        foreach (Shape shape in loadedDoc.GetChildNodes(NodeType.Shape, true))
        {
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue; // Not an OLE object.

            // Retrieve the raw binary data of the OLE object.
            byte[] rawData = oleFormat.GetRawData();

            // Determine a file extension for the extracted data.
            // If the OLE object suggests an extension, use it; otherwise default to .bin.
            string extension = oleFormat.SuggestedExtension ?? ".bin";

            // Build a temporary file name for the extracted data.
            string tempFileName = Path.Combine(Path.GetTempPath(),
                $"ExtractedOle_{Guid.NewGuid()}{extension}");

            // Write the raw data to the temporary file.
            File.WriteAllBytes(tempFileName, rawData);

            // Output the location of the extracted file.
            Console.WriteLine($"OLE object extracted to: {tempFileName}");
        }

        // Clean up the temporary source file.
        if (File.Exists(sourceFilePath))
            File.Delete(sourceFilePath);
    }
}
