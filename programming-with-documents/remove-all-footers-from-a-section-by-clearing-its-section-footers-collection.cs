using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some body text.
        builder.Writeln("Hello World!");

        // Add a primary footer to the first (and only) section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("This is a footer.");

        // Get the first section.
        Section section = doc.Sections[0];

        // Helper to count footers in a section.
        int CountFooters(Section sec) =>
            sec.HeadersFooters
               .OfType<HeaderFooter>()
               .Count(hf =>
                   hf.HeaderFooterType == HeaderFooterType.FooterPrimary ||
                   hf.HeaderFooterType == HeaderFooterType.FooterFirst ||
                   hf.HeaderFooterType == HeaderFooterType.FooterEven);

        // Verify footers exist before removal.
        Console.WriteLine($"Footers before removal: {CountFooters(section)}");

        // Remove all footers from the section.
        List<HeaderFooter> footersToRemove = section.HeadersFooters
            .OfType<HeaderFooter>()
            .Where(hf =>
                hf.HeaderFooterType == HeaderFooterType.FooterPrimary ||
                hf.HeaderFooterType == HeaderFooterType.FooterFirst ||
                hf.HeaderFooterType == HeaderFooterType.FooterEven)
            .ToList();

        foreach (HeaderFooter hf in footersToRemove)
        {
            section.HeadersFooters.Remove(hf);
        }

        // Verify footers have been removed.
        Console.WriteLine($"Footers after removal: {CountFooters(section)}");

        // Save the modified document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
