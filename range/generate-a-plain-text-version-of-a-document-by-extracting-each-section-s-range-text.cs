using System;
using System.IO;
using System.Text;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document with multiple sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Section 1
        builder.Writeln("This is the first section.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2
        builder.Writeln("This is the second section.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 3
        builder.Writeln("This is the third section.");

        // Save the source document locally.
        string sourcePath = "Sample.docx";
        doc.Save(sourcePath);

        // Load the document from the saved file.
        Document loadedDoc = new Document(sourcePath);

        // Extract plain text from each section's Range.Text.
        StringBuilder plainTextBuilder = new StringBuilder();

        foreach (Section section in loadedDoc.Sections)
        {
            // Trim to remove leading/trailing whitespace that may be added by Aspose.Words.
            string sectionText = section.Range.Text.Trim();
            plainTextBuilder.AppendLine(sectionText);
        }

        // Write the combined plain‑text output to a .txt file.
        string outputPath = "PlainTextOutput.txt";
        File.WriteAllText(outputPath, plainTextBuilder.ToString());

        // Optionally, write to console to show completion.
        Console.WriteLine("Plain‑text extraction completed. Output saved to: " + Path.GetFullPath(outputPath));
    }
}
