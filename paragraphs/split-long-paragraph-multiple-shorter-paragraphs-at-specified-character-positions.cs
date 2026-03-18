using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Replacing;

class SplitParagraphExample
{
    static void Main()
    {
        // Create a new document and add a long paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        string longText = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                          "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.";
        builder.Writeln(longText); // single paragraph

        // Positions (character index) where we want to split the paragraph.
        // Example: split after 70 and 150 characters.
        List<int> splitPositions = new List<int> { 70, 150 };

        // Get the first (and only) paragraph.
        Paragraph para = doc.FirstSection.Body.FirstParagraph;

        // Assume the paragraph contains a single Run with the whole text.
        Run run = (Run)para.FirstChild;
        string text = run.Text;

        // Insert a unique marker (U+0001) at the split positions.
        // Process positions in descending order so earlier inserts do not affect later indices.
        char marker = '\u0001';
        splitPositions.Sort((a, b) => b.CompareTo(a));
        foreach (int pos in splitPositions)
        {
            if (pos < 0 || pos > text.Length)
                continue; // ignore invalid positions
            text = text.Insert(pos, marker.ToString());
        }

        // Replace the original run text with the marked text.
        run.Text = text;

        // Use Range.Replace to replace the marker with a paragraph break meta‑character (&p).
        // The meta‑character is interpreted by Aspose.Words during the replace operation.
        FindReplaceOptions options = new FindReplaceOptions();
        doc.Range.Replace(marker.ToString(), "&p", options);

        // Save the result.
        doc.Save("SplitParagraph.docx");
    }
}
