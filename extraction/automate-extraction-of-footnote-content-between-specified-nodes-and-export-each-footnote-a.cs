using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;

public class FootnoteExtractor
{
    public static void Main()
    {
        // Create a new document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a start marker paragraph.
        builder.Writeln("=== START MARKER ===");
        Paragraph startMarker = builder.CurrentParagraph;

        // Add several paragraphs with footnotes between the markers.
        builder.Writeln("First paragraph with a footnote.");
        Footnote fn1 = builder.InsertFootnote(FootnoteType.Footnote, "Footnote text for first paragraph.");

        builder.Writeln("Second paragraph without footnote.");

        builder.Writeln("Third paragraph with two footnotes.");
        Footnote fn2 = builder.InsertFootnote(FootnoteType.Footnote, "First footnote in third paragraph.");
        // Move the builder inside the footnote to add more text.
        builder.MoveTo(fn2.FirstParagraph);
        builder.Writeln("Additional text in the same footnote.");
        // Return to the main document flow.
        builder.MoveToDocumentEnd();

        Footnote fn3 = builder.InsertFootnote(FootnoteType.Footnote, "Second footnote in third paragraph.");

        // Insert an end marker paragraph.
        builder.Writeln("=== END MARKER ===");
        Paragraph endMarker = builder.CurrentParagraph;

        // -----------------------------------------------------------------
        // Extraction: find all footnotes whose reference appears between the
        // start and end marker paragraphs (inclusive) and write each to a file.
        // -----------------------------------------------------------------

        // Get a flat list of all paragraphs in the document.
        NodeCollection allParagraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        // Determine the indexes of the marker paragraphs.
        int startIndex = allParagraphs.IndexOf(startMarker);
        int endIndex = allParagraphs.IndexOf(endMarker);

        if (startIndex < 0 || endIndex < 0 || startIndex > endIndex)
            throw new InvalidOperationException("Invalid marker paragraph positions.");

        int footnoteCounter = 0;

        // Iterate through the paragraphs that lie between the markers.
        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph para = (Paragraph)allParagraphs[i];

            // Collect any footnote nodes that are children of this paragraph.
            NodeCollection footnotesInParagraph = para.GetChildNodes(NodeType.Footnote, true);
            foreach (Footnote footnote in footnotesInParagraph)
            {
                // Extract the footnote's full text.
                string footnoteText = footnote.GetText().Trim();

                // Create a deterministic file name.
                string fileName = $"Footnote_{++footnoteCounter}.txt";

                // Write the footnote text to the file.
                File.WriteAllText(fileName, footnoteText);
            }
        }

        // Validate that at least one footnote file was created.
        if (footnoteCounter == 0)
            throw new InvalidOperationException("No footnotes were found between the specified markers.");

        // Optional: Save the sample document for reference.
        doc.Save("SampleDocument.docx");
    }
}
