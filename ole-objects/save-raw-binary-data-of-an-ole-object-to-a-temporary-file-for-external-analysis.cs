using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare some dummy data to embed as an OLE Package.
        byte[] dummyData = System.Text.Encoding.UTF8.GetBytes("Sample OLE object binary data");

        // Insert the dummy data as an OLE object (Package) into the document.
        using (MemoryStream dataStream = new MemoryStream(dummyData))
        {
            // Insert as an icon (true) – the actual appearance is not important for this example.
            Shape oleShape = builder.InsertOleObject(dataStream, "Package", true, null);

            // Access the OleFormat of the inserted shape.
            OleFormat oleFormat = oleShape.OleFormat;

            // Retrieve the raw binary data of the OLE object.
            byte[] rawOleData = oleFormat.GetRawData();

            // Determine a temporary file path.
            string tempFilePath = Path.Combine(Path.GetTempPath(), "OleObjectData.bin");

            // Write the raw data to the temporary file.
            File.WriteAllBytes(tempFilePath, rawOleData);

            // Optionally, save the document itself (demonstrating the use of the save rule).
            string docPath = Path.Combine(Path.GetTempPath(), "DocumentWithOle.docx");
            doc.Save(docPath);
        }

        // The program finishes without waiting for user input.
    }
}
