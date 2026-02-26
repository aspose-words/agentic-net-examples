using System;
using System.Collections.Generic;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOT (template) file.
        Document doc = new Document("Template.dot");

        // Access the collection of paragraphs in the main body of the first section.
        ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

        // Store the plain text of each paragraph (without the trailing break characters).
        List<string> paragraphTexts = new List<string>();
        foreach (Paragraph para in paragraphs)
        {
            string text = para.GetText();
            // Trim paragraph, cell and section break characters.
            text = text.TrimEnd('\r', '\a', '\f');
            paragraphTexts.Add(text);
        }

        // Example: extract and display the content that lies between the first and last paragraph.
        if (paragraphTexts.Count > 2)
        {
            Console.WriteLine("Content between first and last paragraph:");
            for (int i = 1; i < paragraphTexts.Count - 1; i++)
            {
                Console.WriteLine(paragraphTexts[i]);
            }
        }
        else
        {
            Console.WriteLine("The document does not contain enough paragraphs to extract intermediate content.");
        }
    }
}
