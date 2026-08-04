using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content and a header to the first section.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move to the primary header of the first section and write some text.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header for Section 1");

        // Return to the body of the first section and add some paragraph text.
        builder.MoveToSection(0);
        builder.Writeln("Content of Section 1");

        // Insert a section break to start a new section (Section 2).
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Add body content to the second section.
        builder.Writeln("Content of Section 2");

        // Copy the header from the previous section (Section 0) to the current section (Section 1).
        Section previousSection = doc.Sections[0];
        HeaderFooter previousHeader = previousSection.HeadersFooters[HeaderFooterType.HeaderPrimary];

        if (previousHeader != null)
        {
            // Clone the header so it can be added to another section.
            HeaderFooter clonedHeader = (HeaderFooter)previousHeader.Clone(true);
            // Add the cloned header to the HeadersFooters collection of the second section.
            doc.Sections[1].HeadersFooters.Add(clonedHeader);
        }

        // Define an output path relative to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "CopyHeaderExample.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
