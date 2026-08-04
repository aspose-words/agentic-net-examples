using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        const string inputPath = "input.docx";
        const string outputPath = "output.docx";
        const string keyword = "Special";
        const string replacement = "Replaced";

        // ---------- Create a sample document ----------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Header in first section that contains the keyword.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln($"This is a {keyword} header.");
        builder.MoveToDocumentEnd();

        // Body content.
        builder.Writeln("Body paragraph 1.");

        // ---------- Add a second section ----------
        // Create a new empty section and add it to the document.
        Section secondSection = new Section(doc);
        doc.Sections.Add(secondSection);

        // Move the builder to the newly added section.
        builder.MoveToSection(doc.Sections.Count - 1);

        // Header in second section that does NOT contain the keyword.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Regular header without keyword.");

        // Return to the main story of the second section to add body text.
        builder.MoveToDocumentEnd();
        builder.Writeln("Body paragraph 2.");

        // Save the source document.
        doc.Save(inputPath);

        // ---------- Load the document and replace in matching headers ----------
        Document loaded = new Document(inputPath);
        int totalReplacements = 0;

        foreach (Section section in loaded.Sections)
        {
            HeaderFooter header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
            if (header != null && header.Range.Text.Contains(keyword))
            {
                int replaced = header.Range.Replace(keyword, replacement, new FindReplaceOptions());
                totalReplacements += replaced;
            }
        }

        if (totalReplacements == 0)
            throw new InvalidOperationException("Expected at least one header replacement.");

        // Save the modified document.
        loaded.Save(outputPath);
    }
}
