using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class OleObjectInsertionDemo
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare a simple text file in memory to be inserted as an OLE package.
        byte[] sampleData = System.Text.Encoding.UTF8.GetBytes("Sample OLE package content");
        using (MemoryStream oleStream = new MemoryStream(sampleData))
        {
            // Insert the OLE object from the stream.
            // Parameters: stream, progId ("Package"), asIcon = false, presentation = null.
            Shape oleShape = builder.InsertOleObject(oleStream, "Package", false, null);

            // Verify that the insertion returned a non‑null Shape.
            if (oleShape == null)
                throw new InvalidOperationException("InsertOleObject returned null.");

            // Verify that the Shape contains a valid OleFormat object.
            OleFormat oleFormat = oleShape.OleFormat;
            if (oleFormat == null)
                throw new InvalidOperationException("OleFormat is null after insertion.");

            // Optional: output some properties to confirm successful insertion.
            Console.WriteLine($"OLE object inserted. IsLink: {oleFormat.IsLink}, OleIcon: {oleFormat.OleIcon}");
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "OleObjectDemo.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
