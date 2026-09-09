using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add text to the first section.
        builder.Writeln("Section 1 - Hello World!");

        // Insert a section break to start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Add text to the second section.
        builder.Writeln("Section 2 - Aspose.Words example.");

        // Save the document locally (optional, demonstrates the save rule).
        const string outputPath = "Sections.docx";
        doc.Save(outputPath);

        // Iterate through each section and extract its plain text via the section's Range.Text property.
        for (int i = 0; i < doc.Sections.Count; i++)
        {
            Section section = doc.Sections[i];
            // Trim removes trailing control characters such as section breaks.
            string plainText = section.Range.Text.Trim();
            Console.WriteLine($"Section {i + 1} text: {plainText}");
        }
    }
}
