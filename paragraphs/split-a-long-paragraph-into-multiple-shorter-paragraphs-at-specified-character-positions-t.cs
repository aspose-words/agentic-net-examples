using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class SplitParagraphExample
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // A long paragraph that we will split.
        string longParagraph = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                               "Sed non risus sit amet elit placerat tincidunt. " +
                               "Praesent vitae ligula at odio ultricies gravida. " +
                               "Curabitur euismod, nisl at convallis commodo, " +
                               "nisi lorem fermentum odio, a interdum sapien odio a odio. " +
                               "Vestibulum ante ipsum primis in faucibus orci luctus et " +
                               "ultrices posuere cubilia curae; Integer non turpis " +
                               "vitae ligula aliquet tincidunt. Donec at sapien " +
                               "ullamcorper, dignissim elit non, aliquet massa.";

        // Insert the long paragraph as a single paragraph.
        builder.Writeln(longParagraph);

        // Retrieve the paragraph node that was just created.
        Paragraph originalParagraph = doc.FirstSection.Body.FirstParagraph;

        // Get the text of the paragraph without the trailing paragraph break character.
        string paragraphText = originalParagraph.GetText().TrimEnd('\r');

        // Define character positions where the paragraph should be split.
        // Positions are zero‑based indices in the original string.
        int[] splitPositions = { 80, 160, 240 };

        // Split the text into parts according to the specified positions.
        List<string> parts = new List<string>();
        int start = 0;
        foreach (int pos in splitPositions)
        {
            // Guard against out‑of‑range positions.
            if (pos > paragraphText.Length) break;
            parts.Add(paragraphText.Substring(start, pos - start));
            start = pos;
        }
        // Add the remaining text after the last split position.
        if (start < paragraphText.Length)
            parts.Add(paragraphText.Substring(start));

        // Remove the original long paragraph.
        originalParagraph.Remove();

        // Insert the new shorter paragraphs.
        foreach (string part in parts)
        {
            builder.Writeln(part);
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "SplitParagraph.docx");
        doc.Save(outputPath);
    }
}
