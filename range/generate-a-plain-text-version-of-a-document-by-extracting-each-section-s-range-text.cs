using System;
using System.IO;
using System.Text;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for temporary files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a sample source document with multiple sections.
        string sourcePath = Path.Combine(artifactsDir, "Source.docx");
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // First section.
        builder.Writeln("Section 1 - First paragraph.");
        builder.Writeln("Section 1 - Second paragraph.");
        builder.InsertBreak(BreakType.SectionBreakContinuous);

        // Second section.
        builder.Writeln("Section 2 - Only paragraph.");

        // Save the source document.
        sourceDoc.Save(sourcePath);

        // Load the document from the saved file.
        Document doc = new Document(sourcePath);

        // Extract plain text from each section's range.
        StringBuilder plainText = new StringBuilder();
        for (int i = 0; i < doc.Sections.Count; i++)
        {
            string sectionText = doc.Sections[i].Range.Text.Trim();
            plainText.AppendLine($"--- Section {i + 1} ---");
            plainText.AppendLine(sectionText);
            plainText.AppendLine();
        }

        // Save the extracted plain‑text version.
        string outputPath = Path.Combine(artifactsDir, "PlainText.txt");
        File.WriteAllText(outputPath, plainText.ToString());

        // Optional: display the output path.
        Console.WriteLine($"Plain‑text document saved to: {outputPath}");
    }
}
