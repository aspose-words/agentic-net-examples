using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and add content to two sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Content for the first section.
        builder.Writeln("First section paragraph 1.");
        builder.Writeln("First section paragraph 2.");

        // Insert a new section break.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Content for the second section.
        builder.Writeln("Second section content.");

        // Save the document locally.
        string filePath = "Sample.docx";
        doc.Save(filePath);

        // Load the document back to demonstrate loading.
        Document loadedDoc = new Document(filePath);

        // Extract and display plain text of each section using Section.Range.Text.
        for (int i = 0; i < loadedDoc.Sections.Count; i++)
        {
            Section section = loadedDoc.Sections[i];
            string sectionText = section.Range.Text.Trim(); // Trim control characters.
            Console.WriteLine($"Section {i + 1} text:");
            Console.WriteLine(sectionText);
            Console.WriteLine("---");
        }
    }
}
