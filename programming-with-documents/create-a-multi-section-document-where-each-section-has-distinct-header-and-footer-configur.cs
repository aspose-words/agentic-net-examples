using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Section 1 ----------
        // Ensure the first section does not link to any previous (there is none).
        Section section1 = doc.Sections[0];
        section1.HeadersFooters.LinkToPrevious(false);

        // Header for Section 1
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Header for Section 1");

        // Footer for Section 1
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Footer for Section 1");

        // Content for Section 1
        builder.MoveToDocumentEnd();
        builder.Writeln("This is the content of Section 1.");

        // Insert a section break (new page) to start Section 2.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // ---------- Section 2 ----------
        // Get the newly created second section and break the link to previous headers/footers.
        Section section2 = doc.Sections[1];
        section2.HeadersFooters.LinkToPrevious(false);

        // Header for Section 2
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Header for Section 2");

        // Footer for Section 2
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Footer for Section 2");

        // Content for Section 2
        builder.MoveToDocumentEnd();
        builder.Writeln("This is the content of Section 2.");

        // Save the document.
        string outputPath = "MultiSectionHeadersFooters.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
