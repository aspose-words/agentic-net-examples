using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with two sections, each having a primary header.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First section header contains the keyword "Special".
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Special Header: OldValue");
        builder.MoveToDocumentEnd(); // Return to the main body.
        builder.Writeln("Body text in first section.");

        // Add a second section.
        Section secondSection = new Section(doc);
        doc.Sections.Add(secondSection);

        // Use a builder to edit the second section.
        DocumentBuilder secBuilder = new DocumentBuilder(doc);
        // Move the builder to the newly added section.
        secBuilder.MoveToSection(doc.Sections.Count - 1);
        // Edit the primary header of the second section.
        secBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        secBuilder.Writeln("Regular Header: OldValue");
        secBuilder.MoveToDocumentEnd(); // Return to the main body of the second section.
        secBuilder.Writeln("Body text in second section.");

        // Save the input document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loadedDoc = new Document(inputPath);
        int totalReplacements = 0;
        const string keyword = "Special";
        const string oldText = "OldValue";
        const string newText = "NewValue";

        // Iterate through all sections and their primary headers.
        foreach (Section section in loadedDoc.Sections)
        {
            HeaderFooter? header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
            if (header != null && header.Range.Text.Contains(keyword, StringComparison.Ordinal))
            {
                // Perform replacement only in headers that contain the keyword.
                int replaced = header.Range.Replace(oldText, newText, new FindReplaceOptions());
                totalReplacements += replaced;
            }
        }

        // Validate that at least one replacement occurred.
        if (totalReplacements == 0)
            throw new InvalidOperationException("Expected at least one replacement in headers containing the keyword.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);
    }
}
