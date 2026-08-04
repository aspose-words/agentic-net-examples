using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the source document and the plain‑text output.
        string sourcePath = "Source.docx";
        string outputPath = "PlainTextVersion.txt";

        // -----------------------------------------------------------------
        // Create a sample document with three sections.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Section 1
        builder.Writeln("This is the first section.");
        builder.InsertBreak(BreakType.SectionBreakContinuous);

        // Section 2
        builder.Writeln("Second section goes here.");
        builder.InsertBreak(BreakType.SectionBreakContinuous);

        // Section 3
        builder.Writeln("Third section content.");

        // Save the sample document (bootstrap step).
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // Load the document from the file system.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        // -----------------------------------------------------------------
        // Extract plain text from each section's Range.
        // -----------------------------------------------------------------
        StringBuilder plainText = new StringBuilder();

        for (int i = 0; i < loadedDoc.Sections.Count; i++)
        {
            string sectionText = loadedDoc.Sections[i].Range.Text.Trim();
            plainText.AppendLine($"--- Section {i + 1} ---");
            plainText.AppendLine(sectionText);
            plainText.AppendLine();
        }

        // -----------------------------------------------------------------
        // Create a new document that will hold the concatenated plain text.
        // -----------------------------------------------------------------
        Document plainDoc = new Document();
        DocumentBuilder plainBuilder = new DocumentBuilder(plainDoc);
        plainBuilder.Writeln(plainText.ToString());

        // Save the result as a plain‑text file.
        plainDoc.Save(outputPath, SaveFormat.Text);
    }
}
