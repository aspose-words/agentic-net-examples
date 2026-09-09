using System;
using System.Collections.Generic;
using Aspose.Words;

public class SplitParagraphExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a single long paragraph.
        string longParagraph = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                               "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                               "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris " +
                               "nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in " +
                               "reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla " +
                               "pariatur. Excepteur sint occaecat cupidatat non proident, sunt in " +
                               "culpa qui officia deserunt mollit anim id est laborum.";
        builder.Writeln(longParagraph); // This creates the initial paragraph.

        // Retrieve the first (and only) paragraph.
        Paragraph originalParagraph = doc.FirstSection.Body.FirstParagraph;

        // Get the paragraph text without the trailing paragraph break character.
        string paragraphText = originalParagraph.GetText().TrimEnd('\r');

        // Define character positions where the paragraph should be split.
        // Positions are zero‑based indexes in the original string.
        int[] splitPositions = { 100, 200, 300 };

        // Split the text into parts according to the specified positions.
        List<string> parts = new List<string>();
        int start = 0;
        foreach (int pos in splitPositions)
        {
            if (pos > start && pos < paragraphText.Length)
            {
                parts.Add(paragraphText.Substring(start, pos - start).Trim());
                start = pos;
            }
        }
        // Add the remaining text after the last split position.
        if (start < paragraphText.Length)
            parts.Add(paragraphText.Substring(start).Trim());

        // Remove the original paragraph from the document.
        originalParagraph.Remove();

        // Insert the new shorter paragraphs at the beginning of the document.
        builder.MoveToDocumentStart();
        foreach (string part in parts)
        {
            builder.Writeln(part);
        }

        // Save the resulting document.
        doc.Save("SplitParagraph.docx");
    }
}
