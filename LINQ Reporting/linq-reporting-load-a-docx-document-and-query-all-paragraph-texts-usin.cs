using System;
using System.Linq;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load an existing DOCX file using the Document(string) constructor.
        Document doc = new Document("input.docx");

        // Retrieve all Paragraph nodes in the document (including those in headers/footers).
        var paragraphTexts = doc.GetChildNodes(NodeType.Paragraph, true)
                                .Cast<Paragraph>()
                                .Select(p => p.GetText().Trim())
                                .ToList();

        // Output each paragraph's text to the console.
        foreach (string text in paragraphTexts)
        {
            Console.WriteLine(text);
        }

        // Demonstrate saving the document (uses the Document.Save(string) method).
        doc.Save("output.docx");
    }
}
