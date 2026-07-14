using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build three sections with distinct text.
        builder.Writeln("This is the first section.");
        builder.InsertBreak(BreakType.SectionBreakContinuous);
        builder.Writeln("This is the second section.");
        builder.InsertBreak(BreakType.SectionBreakContinuous);
        builder.Writeln("This is the third section.");

        // Save the sample document (required by the lifecycle rule).
        string docPath = Path.Combine(Environment.CurrentDirectory, "Sample.docx");
        doc.Save(docPath);

        // Extract plain‑text from each section's Range.
        StringBuilder plainTextBuilder = new StringBuilder();

        for (int i = 0; i < doc.Sections.Count; i++)
        {
            // Get the text of the current section.
            string sectionText = doc.Sections[i].Range.Text.Trim();

            // Append the section text to the result, separating sections with a line break.
            plainTextBuilder.AppendLine($"--- Section {i + 1} ---");
            plainTextBuilder.AppendLine(sectionText);
            plainTextBuilder.AppendLine();
        }

        // Write the combined plain‑text to a .txt file.
        string txtPath = Path.Combine(Environment.CurrentDirectory, "PlainTextOutput.txt");
        File.WriteAllText(txtPath, plainTextBuilder.ToString());

        // Optional: display the output path.
        Console.WriteLine($"Plain‑text version saved to: {txtPath}");
    }
}
