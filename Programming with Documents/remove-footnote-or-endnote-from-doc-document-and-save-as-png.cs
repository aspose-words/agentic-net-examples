using Aspose.Words;
using Aspose.Words.Saving;
using System;

class Program
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("Input.docx");

        // All footnotes and endnotes are represented by the Footnote class.
        // Get every Footnote node (this collection contains both footnotes and endnotes).
        NodeCollection footnotes = doc.GetChildNodes(NodeType.Footnote, true);

        // Remove nodes while iterating backwards to avoid collection modification issues.
        for (int i = footnotes.Count - 1; i >= 0; i--)
        {
            footnotes[i].Remove();
        }

        // Save the modified document as a PNG image (first page only).
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // Render only the first page.
            PageSet = new PageSet(0)
        };
        doc.Save("Result.png", pngOptions);
    }
}
