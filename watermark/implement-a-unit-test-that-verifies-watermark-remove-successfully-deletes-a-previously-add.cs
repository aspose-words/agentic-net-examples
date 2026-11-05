using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class WatermarkRemovalTest
{
    // Checks whether the document contains a shape that represents a watermark.
    private static bool HasWatermark(Document doc)
    {
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Aspose.Words typically sets the watermark shape's Name or AlternativeText to contain "Watermark".
            if (!string.IsNullOrEmpty(shape.Name) && shape.Name.Contains("Watermark", StringComparison.OrdinalIgnoreCase))
                return true;

            if (!string.IsNullOrEmpty(shape.AlternativeText) && shape.AlternativeText.Contains("Watermark", StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
    }

    public static void Main()
    {
        // Create a new empty document and add some content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample content for watermark test.");

        // Add a text watermark using the native API.
        doc.Watermark.SetText("CONFIDENTIAL");

        // Verify that the watermark was added.
        if (!HasWatermark(doc))
        {
            Console.WriteLine("Failed: Watermark was not added.");
            Environment.Exit(1);
        }

        // Remove the watermark.
        doc.Watermark.Remove();

        // Verify that the watermark was removed.
        if (HasWatermark(doc))
        {
            Console.WriteLine("Failed: Watermark was not removed.");
            Environment.Exit(1);
        }

        // Save the resulting document to confirm output generation.
        string outputPath = Path.Combine(Path.GetTempPath(), "WatermarkRemovalTest.docx");
        doc.Save(outputPath);

        // Confirm the file exists.
        if (!File.Exists(outputPath))
        {
            Console.WriteLine("Failed: Output file was not saved.");
            Environment.Exit(1);
        }

        Console.WriteLine("Success: Watermark was added, removed, and document saved.");
    }
}
