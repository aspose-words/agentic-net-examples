using System;
using System.Linq;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load an existing DOCM file (uses the Document(string) constructor rule)
        Document doc = new Document("Input.docm");

        // Get all paragraphs in the document as an IEnumerable<Paragraph>
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true)
                            .Cast<Paragraph>();

        // Find paragraphs that contain the placeholder "{Name}"
        var placeholders = paragraphs.Where(p => p.GetText().Contains("{Name}"));

        // Replace the placeholder with actual text in each matching paragraph
        foreach (Paragraph para in placeholders)
        {
            // Range.Replace works on the paragraph's range
            para.Range.Replace("{Name}", "John Doe");
        }

        // Save the modified document (uses the Document.Save(string) rule)
        doc.Save("Output.docx");
    }
}
