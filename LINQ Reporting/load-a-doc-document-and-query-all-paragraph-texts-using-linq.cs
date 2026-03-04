using System;
using System.Linq;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load an existing DOC document from the file system.
        // This uses the Document(string) constructor, which is the provided load rule.
        Document doc = new Document("InputDocument.docx");

        // Retrieve all Paragraph nodes in the document (including those in headers/footers if needed).
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true)
                           .Cast<Paragraph>()
                           .Select(p => p.GetText().Trim())
                           .ToList();

        // Output each paragraph's text.
        foreach (var text in paragraphs)
        {
            Console.WriteLine(text);
        }

        // (Optional) Save the document if any modifications were made.
        // This uses the Document.Save(string) method, which is the provided save rule.
        // doc.Save("OutputDocument.docx");
    }
}
