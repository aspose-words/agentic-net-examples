using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    private static void Assert(bool condition, string message)
    {
        if (!condition)
            throw new Exception("Assertion failed: " + message);
    }

    // Counts shapes that Aspose.Words uses for watermarks (they are named "Watermark").
    private static int CountWatermarkShapes(Document doc)
    {
        int count = 0;
        foreach (Section section in doc.Sections)
        {
            foreach (HeaderFooter headerFooter in section.HeadersFooters)
            {
                NodeCollection shapes = headerFooter.GetChildNodes(NodeType.Shape, true);
                foreach (Shape shape in shapes)
                {
                    if (shape.Name != null && shape.Name.Equals("Watermark", StringComparison.OrdinalIgnoreCase))
                        count++;
                }
            }
        }
        return count;
    }

    public static void Main()
    {
        // Create a temporary folder for test files
        string tempFolder = Path.Combine(Path.GetTempPath(), "WatermarkTest_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempFolder);

        string withWatermarkPath = Path.Combine(tempFolder, "with_watermark.docx");
        string withoutWatermarkPath = Path.Combine(tempFolder, "without_watermark.docx");

        try
        {
            // 1. Create a new empty document
            Document doc = new Document();

            // 2. Add a text watermark
            doc.Watermark.SetText("CONFIDENTIAL");
            doc.Save(withWatermarkPath);

            // Verify that the watermark exists
            int watermarkCountAfterAdd = CountWatermarkShapes(doc);
            Assert(watermarkCountAfterAdd > 0, "Watermark should exist after SetText.");

            // 3. Remove the watermark
            doc.Watermark.Remove();
            doc.Save(withoutWatermarkPath);

            // Verify that the watermark has been removed
            int watermarkCountAfterRemove = CountWatermarkShapes(doc);
            Assert(watermarkCountAfterRemove == 0, "Watermark should be removed after Remove().");

            Console.WriteLine("Watermark removal test passed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Watermark removal test failed: " + ex.Message);
        }
        finally
        {
            // Clean up temporary files
            try { if (File.Exists(withWatermarkPath)) File.Delete(withWatermarkPath); } catch { }
            try { if (File.Exists(withoutWatermarkPath)) File.Delete(withoutWatermarkPath); } catch { }
            try { Directory.Delete(tempFolder, true); } catch { }
        }
    }
}
