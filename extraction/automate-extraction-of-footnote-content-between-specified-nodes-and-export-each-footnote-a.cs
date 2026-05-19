using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // ------------------------------------------------------------
        // 1. Create a sample document with paragraphs and footnotes.
        // ------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Paragraph 0 - introductory text.");

        builder.Writeln("Paragraph 1 - contains a footnote.");
        // Insert a footnote attached to the current paragraph.
        builder.InsertFootnote(FootnoteType.Footnote, "First footnote content.");

        builder.Writeln("Paragraph 2 - another footnote follows.");
        builder.InsertFootnote(FootnoteType.Footnote, "Second footnote content.");

        builder.Writeln("Paragraph 3 - concluding remarks.");

        // Save the source document locally.
        const string sourcePath = "footnote-sample.docx";
        doc.Save(sourcePath);

        // ------------------------------------------------------------
        // 2. Load the document for extraction.
        // ------------------------------------------------------------
        Document loaded = new Document(sourcePath);

        // ------------------------------------------------------------
        // 3. Export each footnote to a separate text file.
        // ------------------------------------------------------------
        int exportedCount = 0;
        foreach (Footnote footnote in loaded.GetChildNodes(NodeType.Footnote, true))
        {
            // Get the plain text of the footnote (trim to remove extra control chars).
            string footnoteText = footnote.GetText().Trim();

            // Create a deterministic file name.
            string fileName = $"footnote-{exportedCount}.txt";

            // Write the footnote text to the file.
            File.WriteAllText(fileName, footnoteText);
            exportedCount++;
        }

        // Validate that at least one footnote file was created.
        if (exportedCount == 0)
            throw new InvalidOperationException("No footnote files were generated.");
    }
}
