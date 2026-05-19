using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class OleInsertionDemo
{
    public static void Main()
    {
        // Create a temporary text file to embed as an OLE object.
        string tempFilePath = Path.Combine(Directory.GetCurrentDirectory(), "SampleData.txt");
        File.WriteAllText(tempFilePath, "This is sample content for OLE embedding.");

        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the OLE object (embedded, not as an icon) using the InsertOleObject overload.
        // The method returns a Shape that contains the OLE object.
        Shape oleShape = builder.InsertOleObject(
            fileName: tempFilePath,
            progId: "Package",   // Generic OLE package.
            isLinked: false,     // Embed the file.
            asIcon: false,       // Display as content, not as an icon.
            presentation: null   // Use default presentation.
        );

        // Verify that the returned Shape reference is not null.
        if (oleShape == null)
        {
            Console.WriteLine("Failed to insert OLE object: returned Shape is null.");
            return;
        }

        // Verify that the Shape contains a valid OleFormat object.
        OleFormat oleFormat = oleShape.OleFormat;
        if (oleFormat == null)
        {
            Console.WriteLine("Inserted Shape does not contain OleFormat.");
            return;
        }

        // Output verification results.
        Console.WriteLine("OLE object inserted successfully.");
        Console.WriteLine($"IsLink: {oleFormat.IsLink}");
        Console.WriteLine($"OleIcon: {oleFormat.OleIcon}");

        // Save the document to the file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OleInsertionResult.docx");
        doc.Save(outputPath);

        // Clean up the temporary file.
        File.Delete(tempFilePath);
    }
}
