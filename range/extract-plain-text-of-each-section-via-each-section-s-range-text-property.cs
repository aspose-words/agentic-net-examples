using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a folder for output files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Path to the sample document.
        string docPath = Path.Combine(artifactsDir, "Sample.docx");

        // Build a document with three sections, each containing its own text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("This is the text of Section 1.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        builder.Writeln("This is the text of Section 2.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        builder.Writeln("This is the text of Section 3.");

        // Save the document so it exists on disk (optional but follows lifecycle rules).
        doc.Save(docPath);

        // Iterate through each section and extract its plain text via the Range.Text property.
        for (int i = 0; i < doc.Sections.Count; i++)
        {
            // The Range.Text includes control characters; Trim removes leading/trailing whitespace.
            string sectionText = doc.Sections[i].Range.Text.Trim();

            Console.WriteLine($"Section {i + 1} text: {sectionText}");
        }
    }
}
