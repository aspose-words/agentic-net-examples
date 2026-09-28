using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    // Checks whether the document contains any watermark shape in its headers/footers.
    private static bool ContainsWatermark(Document doc)
    {
        foreach (Section section in doc.Sections)
        {
            foreach (HeaderFooter headerFooter in section.HeadersFooters)
            {
                // Get all Shape nodes inside the header/footer.
                NodeCollection shapes = headerFooter.GetChildNodes(NodeType.Shape, true);
                foreach (Shape shape in shapes)
                {
                    // A watermark added via Document.Watermark.SetText is a Shape whose TextPath.Text holds the watermark text.
                    // Some versions may leave the Name empty, so we also check the TextPath.Text.
                    if (!string.IsNullOrEmpty(shape.TextPath?.Text))
                        return true;
                }
            }
        }
        return false;
    }

    public static void Main()
    {
        // Paths for intermediate results.
        string watermarkedPath = "watermarked.docx";
        string removedPath = "removed.docx";

        // Create a new blank document and add some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample content for watermark test.");

        // Add a text watermark.
        doc.Watermark.SetText("Test Watermark");
        doc.Save(watermarkedPath);

        // Verify that the watermark was added.
        if (!ContainsWatermark(doc))
            throw new Exception("Failed to add the text watermark.");

        // Remove the watermark.
        doc.Watermark.Remove();
        doc.Save(removedPath);

        // Verify that the watermark was removed.
        if (ContainsWatermark(doc))
            throw new Exception("Failed to remove the text watermark.");

        Console.WriteLine("Watermark removal test passed.");
    }
}
