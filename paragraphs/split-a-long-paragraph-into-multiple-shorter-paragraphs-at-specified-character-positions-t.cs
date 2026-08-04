using System;
using System.Collections.Generic;
using Aspose.Words;

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

        // Retrieve the first (and only) paragraph.
        Paragraph originalParagraph = doc.FirstSection.Body.FirstParagraph;

        // Get the paragraph text without the trailing paragraph break character.
        string paragraphText = originalParagraph.GetText();
        // The paragraph break is the last character (ControlChar.ParagraphBreakChar).
        if (paragraphText.Length > 0 && paragraphText[paragraphText.Length - 1] == ControlChar.ParagraphBreakChar)
            paragraphText = paragraphText.Substring(0, paragraphText.Length - 1);

        // Define split positions (character indices) where new paragraphs should start.
        // Positions are zero‑based and refer to the original text.
        int[] splitPositions = new int[] { 100, 200, 300 };

        // Ensure split positions are within the text length and sorted.
        List<int> positions = new List<int>();
        foreach (int pos in splitPositions)
        {
            if (pos > 0 && pos < paragraphText.Length)
                positions.Add(pos);
        }
        positions.Sort();

        // Build the list of paragraph fragments.
        List<string> fragments = new List<string>();
        int start = 0;
        foreach (int pos in positions)
        {
            fragments.Add(paragraphText.Substring(start, pos - start).Trim());
            start = pos;
        }
        // Add the remaining part.
        fragments.Add(paragraphText.Substring(start).Trim());

        // Remove the original paragraph from the document.
        originalParagraph.Remove();

        // Insert new paragraphs containing the fragments.
        Body body = doc.FirstSection.Body;
        foreach (string fragment in fragments)
        {
            Paragraph newPara = new Paragraph(doc);
            Run run = new Run(doc, fragment);
            newPara.AppendChild(run);
            body.AppendChild(newPara);
        }

        // Save the resulting document.
        doc.Save("SplitParagraph.docx");
    }
}
