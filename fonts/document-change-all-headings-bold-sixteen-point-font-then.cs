using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Path to the source document.
        string inputPath = "input.docx";

        // Path where the modified document will be saved.
        string outputPath = "output.docx";

        // Load the existing document using the Document(string) constructor.
        Document doc = new Document(inputPath);

        // Iterate over all paragraphs in the document.
        foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // Check if the paragraph is a heading (any heading level).
            if (paragraph.ParagraphFormat.IsHeading)
            {
                // Apply bold 16‑point font to every Run within the heading.
                foreach (Run run in paragraph.GetChildNodes(NodeType.Run, true))
                {
                    run.Font.Bold = true;
                    run.Font.Size = 16;
                }
            }
        }

        // Save the modified document using the Document.Save(string) method.
        doc.Save(outputPath);
    }
}
