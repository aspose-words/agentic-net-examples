using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with three sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("This is the first section.");
        builder.InsertBreak(BreakType.SectionBreakContinuous);
        builder.Writeln("This is the second section.");
        builder.InsertBreak(BreakType.SectionBreakContinuous);
        builder.Writeln("This is the third section.");

        // Save the sample document.
        string docPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.docx");
        doc.Save(docPath);

        // Load the document (optional, demonstrates loading workflow).
        Document loadedDoc = new Document(docPath);

        // Extract plain text from each section's range.
        StringBuilder plainTextBuilder = new StringBuilder();
        for (int i = 0; i < loadedDoc.Sections.Count; i++)
        {
            string sectionText = loadedDoc.Sections[i].Range.Text.Trim();
            plainTextBuilder.AppendLine($"--- Section {i + 1} ---");
            plainTextBuilder.AppendLine(sectionText);
            plainTextBuilder.AppendLine();
        }

        // Save the extracted plain‑text version.
        string txtPath = Path.Combine(Directory.GetCurrentDirectory(), "PlainText.txt");
        File.WriteAllText(txtPath, plainTextBuilder.ToString());

        // Optionally, write a confirmation to the console.
        Console.WriteLine($"Plain‑text document saved to: {txtPath}");
    }
}
