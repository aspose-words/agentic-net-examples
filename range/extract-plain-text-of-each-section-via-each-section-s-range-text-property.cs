using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add content to the first section.
        builder.Writeln("Section 1 - First paragraph.");
        builder.Writeln("Section 1 - Second paragraph.");

        // Start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2 - Only paragraph.");

        // Start another new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 3 - First line.");
        builder.Writeln("Section 3 - Second line.");

        // Save the document (optional, demonstrates lifecycle usage).
        const string outputPath = "ExtractSections.docx";
        doc.Save(outputPath);

        // Extract and display the plain text of each section using Section.Range.Text.
        for (int i = 0; i < doc.Sections.Count; i++)
        {
            Section section = doc.Sections[i];
            string sectionText = section.Range.Text.Trim(); // plain text of the section
            Console.WriteLine($"--- Section {i + 1} ---");
            Console.WriteLine(sectionText);
        }
    }
}
