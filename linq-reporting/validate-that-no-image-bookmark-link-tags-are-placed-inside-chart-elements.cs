using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Saving;

class ValidateChartContent
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Iterate over all Shape nodes in the document.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            // Process only shapes that contain a chart.
            if (shape.HasChart)
            {
                // 1. Ensure the shape does not contain an image.
                if (shape.HasImage)
                {
                    Console.WriteLine($"Violation: Chart shape '{shape.Name}' contains an image.");
                }

                // 2. Ensure the shape does not have a hyperlink (HRef property).
                if (!string.IsNullOrEmpty(shape.HRef))
                {
                    Console.WriteLine($"Violation: Chart shape '{shape.Name}' contains a hyperlink (HRef = '{shape.HRef}').");
                }

                // 3. Ensure the shape does not contain any bookmark nodes.
                // Bookmarks are represented by BookmarkStart/BookmarkEnd nodes.
                int bookmarkCount = shape.GetChildNodes(NodeType.BookmarkStart, true).Count;
                if (bookmarkCount > 0)
                {
                    Console.WriteLine($"Violation: Chart shape '{shape.Name}' contains {bookmarkCount} bookmark(s).");
                }
            }
        }

        // Save the (unchanged) document – this step follows the required lifecycle.
        doc.Save("Output.docx", SaveFormat.Docx);
    }
}
