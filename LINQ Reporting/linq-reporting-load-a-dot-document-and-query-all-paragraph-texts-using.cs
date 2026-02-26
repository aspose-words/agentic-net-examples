using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class LinqParagraphsExample
{
    static void Main()
    {
        // Path to the DOT (Word template) file.
        string templatePath = @"C:\Docs\Template.dot";

        // Load the DOT document using the Aspose.Words Document constructor (load rule).
        Document doc = new Document(templatePath);

        // Retrieve all Paragraph nodes in the document (including those inside tables, headers, etc.).
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true)
                           .Cast<Paragraph>()
                           .Select(p => p.GetText().Trim())
                           .Where(text => !string.IsNullOrEmpty(text))
                           .ToList();

        // Output each paragraph text to the console.
        Console.WriteLine("Paragraphs found in the template:");
        foreach (var text in paragraphs)
        {
            Console.WriteLine("- " + text);
        }

        // (Optional) Save the document after processing if you need to persist changes.
        // string outputPath = @"C:\Docs\ProcessedTemplate.docx";
        // doc.Save(outputPath);
    }
}
