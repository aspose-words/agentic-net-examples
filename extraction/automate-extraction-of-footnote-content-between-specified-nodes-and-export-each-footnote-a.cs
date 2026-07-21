using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // -------------------------------------------------
        // Step 1: Create a sample document with footnotes.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Paragraph 1 (no footnote)
        builder.Writeln("Paragraph 1: Introduction.");

        // Paragraph 2 (contains footnote 1)
        // Write the paragraph text without ending the paragraph,
        // insert the footnote, then finish the paragraph.
        builder.Write("Paragraph 2: This sentence has a footnote.");
        builder.InsertFootnote(FootnoteType.Footnote, "First footnote content.");
        builder.Writeln(); // end of paragraph 2

        // Paragraph 3 (contains footnote 2)
        builder.Write("Paragraph 3: Another footnote appears here.");
        builder.InsertFootnote(FootnoteType.Footnote, "Second footnote content.");
        builder.Writeln(); // end of paragraph 3

        // Paragraph 4 (no footnote)
        builder.Writeln("Paragraph 4: Conclusion.");

        // Save the sample document locally.
        const string sourceFile = "source.docx";
        doc.Save(sourceFile);

        // -------------------------------------------------
        // Step 2: Load the document for extraction.
        // -------------------------------------------------
        Document loaded = new Document(sourceFile);

        // Define the start and end paragraph indices (zero‑based).
        // We want to extract footnotes from Paragraph 2 (index 1) to Paragraph 3 (index 2) inclusive.
        int startIndex = 1;
        int endIndex = 2;

        if (loaded.FirstSection.Body.Paragraphs.Count <= endIndex)
            throw new InvalidOperationException("Document does not contain enough paragraphs for the specified range.");

        // -------------------------------------------------
        // Step 3: Collect footnotes that belong to the selected paragraphs.
        // -------------------------------------------------
        var footnotes = new List<Footnote>();
        var seenFootnoteIds = new HashSet<int>();

        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph para = loaded.FirstSection.Body.Paragraphs[i];
            // Footnote nodes are inline children of the paragraph.
            NodeCollection footnoteNodes = para.GetChildNodes(NodeType.Footnote, true);
            foreach (Node node in footnoteNodes)
            {
                if (node is Footnote fn)
                {
                    int id = fn.GetHashCode(); // unique identifier for this run
                    if (!seenFootnoteIds.Contains(id))
                    {
                        footnotes.Add(fn);
                        seenFootnoteIds.Add(id);
                    }
                }
            }
        }

        // -------------------------------------------------
        // Step 4: Export each collected footnote to a separate text file.
        // -------------------------------------------------
        int fileIndex = 0;
        foreach (Footnote fn in footnotes)
        {
            string fileName = $"footnote-{fileIndex}.txt";
            // GetText returns the footnote marker plus its content; Trim removes extra whitespace.
            File.WriteAllText(fileName, fn.GetText().Trim());
            fileIndex++;
        }

        // Validation: ensure at least one footnote file was created.
        if (fileIndex == 0)
            throw new InvalidOperationException("No footnote files were generated.");

        // Optional cleanup of the sample document.
        // File.Delete(sourceFile);
    }
}
