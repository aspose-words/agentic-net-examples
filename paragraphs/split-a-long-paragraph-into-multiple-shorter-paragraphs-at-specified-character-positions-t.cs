using System;
using System.Collections.Generic;
using Aspose.Words;

public class SplitParagraphExample
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a long paragraph that we will split later.
        string longText = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                          "Sed non risus. Suspendisse lectus tortor, dignissim sit " +
                          "amet, adipiscing nec, ultricies sed, dolor. Cras elementum " +
                          "ultrices diam. Maecenas ligula massa, varius a, semper congue, " +
                          "euismod non, mi. Proin porttitor, orci nec nonummy molestie, " +
                          "enim est eleifend mi, non fermentum diam nisl sit amet erat. " +
                          "Duis semper. Duis arcu massa, scelerisque vitae, consequat in, " +
                          "pretium a, enim. Pellentesque congue.";
        builder.Writeln(longText);

        // Define character positions where the paragraph should be split.
        // Positions are zero‑based indices in the paragraph text (excluding the paragraph mark).
        int[] splitPositions = new int[] { 80, 160, 240 };

        // Retrieve the first (and only) paragraph.
        Body body = doc.FirstSection.Body;
        Paragraph originalParagraph = body.FirstParagraph;

        // Get the paragraph text without the terminating paragraph mark.
        string paragraphText = originalParagraph.GetText().TrimEnd('\r', '\n');

        // Ensure split positions are sorted and within the text length.
        Array.Sort(splitPositions);
        List<int> validSplits = new List<int>();
        foreach (int pos in splitPositions)
        {
            if (pos > 0 && pos < paragraphText.Length)
                validSplits.Add(pos);
        }

        // Split the text according to the positions.
        List<string> parts = new List<string>();
        int previousIndex = 0;
        foreach (int split in validSplits)
        {
            parts.Add(paragraphText.Substring(previousIndex, split - previousIndex));
            previousIndex = split;
        }
        // Add the remaining part.
        parts.Add(paragraphText.Substring(previousIndex));

        // Insert new paragraphs after the original one.
        Paragraph insertAfter = originalParagraph;
        foreach (string part in parts)
        {
            Paragraph newParagraph = new Paragraph(doc);
            newParagraph.AppendChild(new Run(doc, part));
            body.InsertAfter(newParagraph, insertAfter);
            insertAfter = newParagraph;
        }

        // Remove the original long paragraph.
        originalParagraph.Remove();

        // Save the resulting document.
        doc.Save("SplitParagraph.docx");
    }
}
