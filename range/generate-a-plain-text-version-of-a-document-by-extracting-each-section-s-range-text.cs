using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document with multiple sections.
        Document doc = new Document();
        doc.Sections.Clear();

        for (int i = 1; i <= 3; i++)
        {
            // Create a new section.
            Section section = new Section(doc);
            // Add a body to the section.
            Body body = new Body(doc);
            section.AppendChild(body);
            // Add a paragraph with some text.
            Paragraph para = new Paragraph(doc);
            Run run = new Run(doc, $"This is the text of section {i}.");
            para.AppendChild(run);
            body.AppendChild(para);
            // Add the section to the document.
            doc.Sections.Add(section);
        }

        // Extract plain text from each section's Range.
        StringWriter plainTextWriter = new StringWriter();
        foreach (Section sec in doc.Sections)
        {
            plainTextWriter.WriteLine(sec.Range.Text.Trim());
        }

        string plainText = plainTextWriter.ToString();

        // Write the extracted text to a file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "PlainTextOutput.txt");
        File.WriteAllText(outputPath, plainText);

        // Also output to console for verification.
        Console.WriteLine("Extracted plain‑text content:");
        Console.WriteLine(plainText);
        Console.WriteLine($"Plain‑text file saved to: {outputPath}");
    }
}
