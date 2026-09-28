using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample long paragraph text.
        string longText = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                          "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                          "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris " +
                          "nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in " +
                          "reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla " +
                          "pariatur. Excepteur sint occaecat cupidatat non proident, sunt in " +
                          "culpa qui officia deserunt mollit anim id est laborum.";

        // Insert the long paragraph into the document.
        builder.Writeln(longText);

        // Retrieve the inserted paragraph.
        Paragraph originalParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);
        string paragraphText = originalParagraph.GetText().TrimEnd('\r', '\n');

        // Define the maximum length of each split segment.
        int maxSegmentLength = 80; // characters

        // Split the paragraph text into segments.
        List<string> segments = new List<string>();
        for (int i = 0; i < paragraphText.Length; i += maxSegmentLength)
        {
            int length = Math.Min(maxSegmentLength, paragraphText.Length - i);
            segments.Add(paragraphText.Substring(i, length));
        }

        // Replace the original paragraph with the first segment.
        originalParagraph.Runs.Clear();
        originalParagraph.AppendChild(new Run(doc, segments[0]));

        // Insert remaining segments as new paragraphs after the original.
        Paragraph previousParagraph = originalParagraph;
        for (int i = 1; i < segments.Count; i++)
        {
            Paragraph newParagraph = new Paragraph(doc);
            newParagraph.AppendChild(new Run(doc, segments[i]));
            previousParagraph.ParentNode.InsertAfter(newParagraph, previousParagraph);
            previousParagraph = newParagraph;
        }

        // Save the resulting document.
        string outputPath = "SplitParagraph.docx";
        doc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Document saved to '{outputPath}'.");
    }
}
