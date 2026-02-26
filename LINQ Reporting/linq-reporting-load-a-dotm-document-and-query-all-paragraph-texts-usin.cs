using System;
using System.Linq;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the DOTM template file.
        string templatePath = "Template.dotm";

        // Load the DOTM document.
        Document doc = new Document(templatePath);

        // Get all Paragraph nodes in the document and extract their text.
        var paragraphTexts = doc.GetChildNodes(NodeType.Paragraph, true)
                               .Cast<Paragraph>()
                               .Select(p => p.GetText().Trim());

        // Output each paragraph's text to the console.
        foreach (string text in paragraphTexts)
        {
            Console.WriteLine(text);
        }
    }
}
