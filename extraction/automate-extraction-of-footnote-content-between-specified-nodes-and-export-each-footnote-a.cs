using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // -------------------------------------------------
        // 1. Create a sample document containing footnotes.
        // -------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        builder.Writeln("Paragraph 0 - introductory text.");

        // Paragraph 1 with a footnote.
        builder.Writeln("Paragraph 1 with a footnote reference.");
        Footnote footnote1 = new Footnote(sampleDoc, FootnoteType.Footnote);
        Paragraph footnotePara1 = new Paragraph(sampleDoc);
        footnotePara1.AppendChild(new Run(sampleDoc, "First footnote content."));
        footnote1.AppendChild(footnotePara1);
        builder.CurrentParagraph.AppendChild(footnote1);

        // Paragraph 2 without footnote.
        builder.Writeln("Paragraph 2 - no footnote here.");

        // Paragraph 3 with another footnote.
        builder.Writeln("Paragraph 3 with a second footnote reference.");
        Footnote footnote2 = new Footnote(sampleDoc, FootnoteType.Footnote);
        Paragraph footnotePara2 = new Paragraph(sampleDoc);
        footnotePara2.AppendChild(new Run(sampleDoc, "Second footnote content."));
        footnote2.AppendChild(footnotePara2);
        builder.CurrentParagraph.AppendChild(footnote2);

        // Paragraph 4 - final.
        builder.Writeln("Paragraph 4 - end of document.");

        // Save the sample document locally.
        const string inputFileName = "footnote-sample.docx";
        sampleDoc.Save(inputFileName);

        // -------------------------------------------------
        // 2. Load the document and define extraction range.
        // -------------------------------------------------
        Document loadedDoc = new Document(inputFileName);
        Body body = loadedDoc.FirstSection.Body;

        // Define start (Paragraph 1) and end (Paragraph 3) inclusive.
        Paragraph startParagraph = body.Paragraphs[1]; // "Paragraph 1 ..."
        Paragraph endParagraph = body.Paragraphs[3];   // "Paragraph 3 ..."
        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Start or end paragraph not found.");

        int startIndex = body.Paragraphs.IndexOf(startParagraph);
        int endIndex = body.Paragraphs.IndexOf(endParagraph);
        if (startIndex < 0 || endIndex < 0 || startIndex > endIndex)
            throw new InvalidOperationException("Invalid paragraph range.");

        // -------------------------------------------------
        // 3. Extract footnotes that appear within the range.
        // -------------------------------------------------
        int fileIndex = 0;
        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph para = body.Paragraphs[i];

            // Footnote nodes act as the reference and also contain the footnote content.
            IEnumerable<Footnote> footnotesInPara = para.GetChildNodes(NodeType.Footnote, true)
                                                       .Cast<Footnote>();

            foreach (Footnote footnote in footnotesInPara)
            {
                string outputFileName = $"footnote-{fileIndex}.txt";
                // GetText returns the footnote's paragraph text plus a trailing line break.
                File.WriteAllText(outputFileName, footnote.GetText().Trim());
                fileIndex++;
            }
        }

        // -------------------------------------------------
        // 4. Validation – ensure at least one file was created.
        // -------------------------------------------------
        if (fileIndex == 0)
            throw new InvalidOperationException("No footnote files were generated.");

        for (int i = 0; i < fileIndex; i++)
        {
            string path = $"footnote-{i}.txt";
            if (!File.Exists(path))
                throw new InvalidOperationException($"Expected output file '{path}' was not created.");
        }

        // Optional: indicate success.
        Console.WriteLine($"{fileIndex} footnote file(s) successfully created.");
    }
}
