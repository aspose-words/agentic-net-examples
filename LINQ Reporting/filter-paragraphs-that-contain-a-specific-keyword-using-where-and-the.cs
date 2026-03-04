using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOCM file
        string inputPath = @"Input.docm";

        // Path to the output DOCX file that will contain only the filtered paragraphs
        string outputPath = @"Filtered.docx";

        // Keyword to search for (case‑insensitive)
        string keyword = "Aspose";

        // Load the DOCM document
        Document doc = new Document(inputPath);

        // Retrieve all paragraphs from the main story of the document
        ParagraphCollection allParagraphs = doc.FirstSection.Body.Paragraphs;

        // Filter paragraphs that contain the keyword using LINQ Where and string Contains (case‑insensitive)
        List<Paragraph> matchingParagraphs = allParagraphs
            .Cast<Paragraph>()
            .Where(p => p.GetText().IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
            .ToList();

        // Create a new empty document to hold the filtered paragraphs
        Document result = new Document();
        result.RemoveAllChildren(); // Remove the default section, body, and paragraph

        // Add a new section and body to the result document
        Section section = new Section(result);
        result.AppendChild(section);
        Body body = new Body(result);
        section.AppendChild(body);

        // Import each matching paragraph into the new document and append it to the body
        foreach (Paragraph para in matchingParagraphs)
        {
            Node imported = result.ImportNode(para, true);
            body.AppendChild(imported);
        }

        // Save the filtered document
        result.Save(outputPath);
    }
}
