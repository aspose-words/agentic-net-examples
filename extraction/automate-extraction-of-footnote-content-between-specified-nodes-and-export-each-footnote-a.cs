using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample document with several paragraphs and footnotes.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Paragraph 0 (no footnote)
        builder.Writeln("Paragraph 0 - introductory text.");

        // Paragraph 1 (footnote A)
        builder.Writeln("Paragraph 1 - contains footnote A.");
        Footnote footnoteA = builder.InsertFootnote(FootnoteType.Footnote, "Footnote A content.");

        // Paragraph 2 (no footnote)
        builder.Writeln("Paragraph 2 - plain text.");

        // Paragraph 3 (footnote B)
        builder.Writeln("Paragraph 3 - contains footnote B.");
        Footnote footnoteB = builder.InsertFootnote(FootnoteType.Footnote, "Footnote B content.");

        // Paragraph 4 (footnote C)
        builder.Writeln("Paragraph 4 - contains footnote C.");
        Footnote footnoteC = builder.InsertFootnote(FootnoteType.Footnote, "Footnote C content.");

        // Save the sample document.
        const string inputPath = "footnote-input.docx";
        doc.Save(inputPath);

        // Load the document for extraction.
        Document loaded = new Document(inputPath);

        // Define the range: from Paragraph 1 to Paragraph 3 inclusive.
        Body body = loaded.FirstSection.Body;
        Paragraph startParagraph = body.Paragraphs[1];
        Paragraph endParagraph = body.Paragraphs[3];

        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Start or end paragraph not found.");

        // Determine the indexes of the boundary paragraphs.
        int startIndex = body.Paragraphs.IndexOf(startParagraph);
        int endIndex = body.Paragraphs.IndexOf(endParagraph);
        if (startIndex < 0 || endIndex < 0 || startIndex > endIndex)
            throw new InvalidOperationException("Invalid paragraph range.");

        // Extract footnotes whose parent paragraph lies within the specified range.
        int footnoteCounter = 0;
        foreach (Footnote footnote in loaded.GetChildNodes(NodeType.Footnote, true))
        {
            Paragraph parentPara = footnote.ParentParagraph;
            if (parentPara == null)
                continue;

            int paraIndex = body.Paragraphs.IndexOf(parentPara);
            if (paraIndex >= startIndex && paraIndex <= endIndex)
            {
                string fileName = $"footnote-{footnoteCounter}.txt";
                File.WriteAllText(fileName, footnote.GetText().Trim());
                footnoteCounter++;
            }
        }

        // Validate that at least one footnote file was generated.
        if (footnoteCounter == 0)
            throw new InvalidOperationException("No footnote files were generated.");

        // Optional: clean up the sample document file (not required).
        // File.Delete(inputPath);
    }
}
