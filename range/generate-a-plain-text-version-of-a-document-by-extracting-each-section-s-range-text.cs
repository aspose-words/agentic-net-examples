using System;
using System.IO;
using System.Text;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file paths.
        string docPath = "Sample.docx";
        string txtPath = "PlainTextOutput.txt";

        // Create a sample document with multiple sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First section.
        builder.Writeln("Section 1: Introduction");
        builder.Writeln("This is the first section.");

        // Insert a continuous section break.
        builder.InsertBreak(BreakType.SectionBreakContinuous);

        // Second section.
        builder.Writeln("Section 2: Details");
        builder.Writeln("This is the second section.");

        // Insert a new page section break.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Third section.
        builder.Writeln("Section 3: Conclusion");
        builder.Writeln("This is the third section.");

        // Save the source document.
        doc.Save(docPath);

        // Load the document from the saved file.
        Document loadedDoc = new Document(docPath);

        // Extract plain text from each section's range.
        StringBuilder plainTextBuilder = new StringBuilder();

        for (int i = 0; i < loadedDoc.Sections.Count; i++)
        {
            var section = loadedDoc.Sections[i];
            string sectionText = section.Range.Text.Trim();

            plainTextBuilder.AppendLine($"--- Section {i + 1} ---");
            plainTextBuilder.AppendLine(sectionText);
            plainTextBuilder.AppendLine();
        }

        // Write the extracted text to a plain‑text file.
        File.WriteAllText(txtPath, plainTextBuilder.ToString());

        // Optionally, display a confirmation.
        Console.WriteLine($"Plain‑text extraction completed. Output saved to '{txtPath}'.");
    }
}
