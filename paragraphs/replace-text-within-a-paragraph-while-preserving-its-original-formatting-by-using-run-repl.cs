using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder for easy content insertion.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with mixed formatting.
        builder.Writeln("This is a sample paragraph.");
        builder.Writeln(); // Add an empty line.

        // Start a new paragraph.
        Paragraph para = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(para);

        // Add runs with different formatting.
        Run run1 = new Run(doc, "Hello ");
        para.AppendChild(run1);

        Run runBold = new Run(doc, "World");
        runBold.Font.Bold = true; // Preserve bold formatting.
        para.AppendChild(runBold);

        Run run2 = new Run(doc, "! This is a test.");
        para.AppendChild(run2);

        // Replace the word "World" with "Universe" while preserving formatting.
        string target = "World";
        string replacement = "Universe";

        foreach (Run run in para.Runs)
        {
            if (run.Text.Contains(target))
            {
                // Preserve the original formatting by only changing the text.
                run.Text = run.Text.Replace(target, replacement);
            }
        }

        // Save the document.
        doc.Save("Output.docx");
    }
}
