using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Section 1 ----------
        // Add a primary header to the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Header for Section 1");

        // Add some body text to the first section.
        builder.MoveToDocumentEnd();
        builder.Writeln("Body of Section 1.");

        // Insert a section break to start Section 2.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // ---------- Section 2 ----------
        // Copy the header from the previous (first) section into the current (second) section.
        HeaderFooter sourceHeader = doc.Sections[0].HeadersFooters[HeaderFooterType.HeaderPrimary];
        HeaderFooter clonedHeader = (HeaderFooter)sourceHeader.Clone(true);
        // Add the cloned header to the current section's header collection.
        doc.Sections[1].HeadersFooters.Add(clonedHeader);

        // Add some body text to the second section.
        builder.Writeln("Body of Section 2.");

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was saved and can be reopened.
        Document reopened = new Document(outputPath);
    }
}
