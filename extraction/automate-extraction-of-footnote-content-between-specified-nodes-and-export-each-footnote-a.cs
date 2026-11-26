using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;

public class FootnoteExtractionExample
{
    public static void Main()
    {
        // Create a sample document with footnotes placed between START and END markers.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Intro paragraph (outside the range).
        builder.Writeln("Intro paragraph before range.");

        // START marker.
        builder.Writeln("START");
        // Paragraphs with footnotes that lie between START and END.
        for (int i = 1; i <= 2; i++)
        {
            builder.Writeln($"Paragraph {i} with a footnote.");
            // Create a footnote.
            Footnote footnote = new Footnote(doc, FootnoteType.Footnote);
            Paragraph footnoteParagraph = new Paragraph(doc);
            footnoteParagraph.AppendChild(new Run(doc, $"Footnote {i} content."));
            footnote.AppendChild(footnoteParagraph);
            // Attach the footnote to the current paragraph.
            builder.CurrentParagraph.AppendChild(footnote);
        }

        // END marker.
        builder.Writeln("END");

        // Paragraph after the range (no footnote).
        builder.Writeln("Paragraph after range.");

        // Save the sample document.
        const string inputPath = "sample.docx";
        doc.Save(inputPath);

        // Load the document for extraction.
        Document loaded = new Document(inputPath);

        // Extract all footnotes (they are all within the defined range).
        int footnoteIndex = 0;
        foreach (Footnote footnote in loaded.GetChildNodes(NodeType.Footnote, true))
        {
            string footnoteText = footnote.GetText().Trim();
            string outputFile = $"footnote-{footnoteIndex}.txt";
            File.WriteAllText(outputFile, footnoteText);
            footnoteIndex++;
        }

        // Validate that at least one footnote file was created.
        if (footnoteIndex == 0)
            throw new InvalidOperationException("No footnote files were generated.");

        // Optional: Verify that the expected number of files exist.
        for (int i = 0; i < footnoteIndex; i++)
        {
            string expectedFile = $"footnote-{i}.txt";
            if (!File.Exists(expectedFile))
                throw new InvalidOperationException($"Expected footnote file '{expectedFile}' was not created.");
        }
    }
}
