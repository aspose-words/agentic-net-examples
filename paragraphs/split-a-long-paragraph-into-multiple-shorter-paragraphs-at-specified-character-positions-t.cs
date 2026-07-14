using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class SplitParagraphExample
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

        // Get the first (and only) paragraph that contains the text.
        Paragraph originalParagraph = doc.FirstSection.Body.FirstParagraph;

        // Retrieve the paragraph text without the trailing paragraph break character.
        string paragraphText = originalParagraph.GetText().TrimEnd('\r');

        // Define character positions where the paragraph should be split.
        // Positions are zero‑based indexes into the string.
        int[] splitPositions = new int[] { 100, 200, 300 };

        // Ensure split positions are sorted and within the text length.
        Array.Sort(splitPositions);
        int textLength = paragraphText.Length;
        for (int i = splitPositions.Length - 1; i >= 0; i--)
        {
            if (splitPositions[i] <= 0 || splitPositions[i] >= textLength)
                splitPositions[i] = -1; // Mark invalid positions.
        }

        // Build the list of substring pieces.
        var pieces = new System.Collections.Generic.List<string>();
        int start = 0;
        foreach (int pos in splitPositions)
        {
            if (pos == -1) continue;
            int length = pos - start;
            if (length > 0)
                pieces.Add(paragraphText.Substring(start, length));
            start = pos;
        }
        // Add the remaining text.
        if (start < textLength)
            pieces.Add(paragraphText.Substring(start));

        // Remove the original paragraph from the document.
        originalParagraph.Remove();

        // Insert the new shorter paragraphs into the body.
        Body body = doc.FirstSection.Body;
        foreach (string piece in pieces)
        {
            Paragraph newPara = new Paragraph(doc);
            Run run = new Run(doc, piece);
            newPara.AppendChild(run);
            body.AppendChild(newPara);
        }

        // Save the resulting document.
        doc.Save("SplitParagraph.docx");
    }
}
