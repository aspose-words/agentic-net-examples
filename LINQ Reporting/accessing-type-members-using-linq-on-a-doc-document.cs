using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load an existing Word document from file.
        Document doc = new Document("Input.docx");

        // -----------------------------------------------------------------
        // Example 1: Use LINQ to find all paragraphs that contain the word "Aspose"
        // and print their text to the console.
        // -----------------------------------------------------------------
        var paragraphsWithAspose = doc
            .GetChildNodes(NodeType.Paragraph, true)          // Get all paragraph nodes in the document.
            .OfType<Paragraph>()                              // Cast to Paragraph type.
            .Where(p => p.GetText().IndexOf("Aspose", StringComparison.OrdinalIgnoreCase) >= 0);

        foreach (var paragraph in paragraphsWithAspose)
        {
            Console.WriteLine("Paragraph containing 'Aspose':");
            Console.WriteLine(paragraph.GetText().Trim());
        }

        // -----------------------------------------------------------------
        // Example 2: Use LINQ to find all runs (pieces of text) that are bold
        // and replace the word "old" with "new" inside those runs.
        // -----------------------------------------------------------------
        var boldRuns = doc
            .GetChildNodes(NodeType.Run, true)               // Get all run nodes.
            .OfType<Run>()                                   // Cast to Run type.
            .Where(r => r.Font.Bold);                        // Filter runs that are bold.

        foreach (var run in boldRuns)
        {
            if (run.Text.Contains("old"))
            {
                run.Text = run.Text.Replace("old", "new");
            }
        }

        // -----------------------------------------------------------------
        // Example 3: Use LINQ to collect all images (Shape nodes) in the document
        // and output their dimensions.
        // -----------------------------------------------------------------
        var images = doc
            .GetChildNodes(NodeType.Shape, true)             // Get all shape nodes.
            .OfType<Shape>()                                 // Cast to Shape type.
            .Where(s => s.IsImage);                          // Filter only image shapes.

        foreach (var image in images)
        {
            Console.WriteLine($"Image found: Width={image.Width}, Height={image.Height}");
        }

        // Save the modified document to a new file.
        doc.Save("Output.docx");
    }
}
