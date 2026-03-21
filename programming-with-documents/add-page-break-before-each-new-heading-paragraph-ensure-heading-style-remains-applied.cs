using System;
using System.IO;
using Aspose.Words;

class AddPageBreakBeforeHeadings
{
    static void Main()
    {
        // Create a new document with some heading paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a normal paragraph.
        builder.Writeln("This is a normal paragraph.");

        // Add Heading 1.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1");

        // Add another normal paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Another normal paragraph.");

        // Add Heading 2.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 2");

        // Iterate through all paragraphs and set PageBreakBefore for headings.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.ParagraphFormat.StyleIdentifier >= StyleIdentifier.Heading1 &&
                para.ParagraphFormat.StyleIdentifier <= StyleIdentifier.Heading9)
            {
                para.ParagraphFormat.PageBreakBefore = true;
            }
        }

        // Save the modified document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Output.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
