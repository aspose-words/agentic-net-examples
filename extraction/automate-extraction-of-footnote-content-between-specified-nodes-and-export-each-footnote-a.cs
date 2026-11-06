using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create a sample document with footnotes.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Paragraph 1 with a footnote.
        builder.Writeln("Paragraph 1.");
        InsertFootnote(builder, "Footnote 1 content.");

        // Paragraph 2 with a footnote.
        builder.Writeln("Paragraph 2.");
        InsertFootnote(builder, "Footnote 2 content.");

        // Paragraph 3 without a footnote.
        builder.Writeln("Paragraph 3.");

        // Paragraph 4 with a footnote.
        builder.Writeln("Paragraph 4.");
        InsertFootnote(builder, "Footnote 4 content.");

        // Save the source document.
        string sourcePath = "footnote-input.docx";
        doc.Save(sourcePath);

        // Load the document for extraction.
        Document loaded = new Document(sourcePath);

        // Define start and end paragraphs (second and fourth paragraphs).
        Paragraph startPara = loaded.FirstSection.Body.Paragraphs[1];
        Paragraph endPara = loaded.FirstSection.Body.Paragraphs[3];

        if (startPara == null || endPara == null)
            throw new InvalidOperationException("Boundary paragraphs not found.");

        // Extract footnotes that belong to paragraphs between the boundaries.
        int footnoteIndex = 0;
        bool withinRange = false;

        foreach (Paragraph para in loaded.FirstSection.Body.Paragraphs)
        {
            if (para == startPara)
                withinRange = true;

            if (withinRange)
            {
                foreach (Node node in para.GetChildNodes(NodeType.Footnote, true))
                {
                    Footnote footnote = node as Footnote;
                    if (footnote != null)
                    {
                        string fileName = $"footnote-{footnoteIndex}.txt";
                        File.WriteAllText(fileName, footnote.GetText().Trim());
                        footnoteIndex++;
                    }
                }
            }

            if (para == endPara)
                break;
        }

        if (footnoteIndex == 0)
            throw new InvalidOperationException("No footnote files were generated.");

        // Verify that each footnote file was created.
        for (int i = 0; i < footnoteIndex; i++)
        {
            string fileName = $"footnote-{i}.txt";
            if (!File.Exists(fileName))
                throw new InvalidOperationException($"Expected footnote file {fileName} was not created.");
        }
    }

    private static void InsertFootnote(DocumentBuilder builder, string text)
    {
        // Create a footnote and attach it to the current paragraph.
        Footnote footnote = new Footnote(builder.Document, FootnoteType.Footnote);
        Paragraph footnoteParagraph = new Paragraph(builder.Document);
        footnoteParagraph.AppendChild(new Run(builder.Document, text));
        footnote.AppendChild(footnoteParagraph);
        builder.CurrentParagraph.AppendChild(footnote);
    }
}
