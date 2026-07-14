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
        builder.Writeln("First section text.");

        // Insert a new page section break and add text to the second section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Second section text.");

        // Insert a continuous section break and add text to the third section.
        builder.InsertBreak(BreakType.SectionBreakContinuous);
        builder.Writeln("Third section text.");

        // Save the document to the local file system.
        const string filePath = "Sections.docx";
        doc.Save(filePath);

        // Iterate through each section and extract its plain text via the Section's Range.Text property.
        for (int i = 0; i < doc.Sections.Count; i++)
        {
            Section section = doc.Sections[i];
            // Trim removes trailing control characters such as section breaks.
            string sectionText = section.Range.Text.Trim();
            Console.WriteLine($"Section {i + 1} text: {sectionText}");
        }
    }
}
