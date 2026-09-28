using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Section 1 ----------
        // Add some body text.
        builder.Writeln("Content of Section 1.");

        // Add a primary header to Section 1.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Header for Section 1");
        builder.MoveToDocumentEnd(); // Return to the main story.

        // Insert a section break to start Section 2.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // ---------- Section 2 ----------
        // Add body text for Section 2.
        builder.Writeln("Content of Section 2.");

        // Copy the header from the previous section (Section 1) into Section 2.
        // Get the previous section (index 0) and its primary header.
        Section previousSection = doc.Sections[0];
        HeaderFooter previousHeader = previousSection.HeadersFooters[HeaderFooterType.HeaderPrimary];

        // Get the current section (index 1).
        Section currentSection = doc.Sections[1];

        // Clone the previous header and add it to the current section's header collection.
        currentSection.HeadersFooters.Add(previousHeader.Clone(true));

        // Save the document to disk.
        string outputPath = "HeaderCopyExample.docx";
        doc.Save(outputPath);
    }
}
