using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First section content.
        builder.Writeln("Section 1 - Line 1");
        builder.Writeln("Section 1 - Line 2");

        // Insert a section break to start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second section content.
        builder.Writeln("Section 2 - Line 1");
        builder.Writeln("Section 2 - Line 2");

        // Insert another section break for a third section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 3 - Only line");

        // Iterate through each section and extract its plain text via the Section's Range.Text property.
        for (int i = 0; i < doc.Sections.Count; i++)
        {
            var section = doc.Sections[i];
            string sectionText = section.Range.Text.Trim(); // Trim to remove trailing control characters.
            Console.WriteLine($"--- Section {i + 1} ---");
            Console.WriteLine(sectionText);
            Console.WriteLine();
        }

        // Save the document to verify the content (optional).
        doc.Save("ExtractedSections.docx");
    }
}
