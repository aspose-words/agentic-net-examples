using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using Aspose.Words.Lists;
using Aspose.Words.Replacing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Define the folder where the output document will be saved.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();

        // Remove any default nodes (the document initially contains a section, body and a paragraph).
        doc.RemoveAllChildren();

        // Add a new section to the document.
        Section section = new Section(doc);
        doc.AppendChild(section);

        // Add a body to the section.
        Body body = new Body(doc);
        section.AppendChild(body);

        // Create a new paragraph.
        Paragraph para = new Paragraph(doc);

        // Apply paragraph formatting:
        // - Use the built‑in "Heading 1" style.
        // - Center the paragraph text.
        // - Add 12 points of space after the paragraph.
        // - Keep the paragraph together on one page.
        para.ParagraphFormat.StyleName = "Heading 1";
        para.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        para.ParagraphFormat.SpaceAfter = 12;
        para.ParagraphFormat.KeepTogether = true;

        // Append the paragraph to the body.
        body.AppendChild(para);

        // Create a run (the actual text) and add it to the paragraph.
        Run run = new Run(doc);
        run.Text = "Formatted paragraph using Aspose.Words.";
        // Optionally set font properties for the run.
        run.Font.Name = "Arial";
        run.Font.Size = 16;
        run.Font.Color = System.Drawing.Color.Blue;
        para.AppendChild(run);

        // Save the document in DOCX format.
        string outputPath = Path.Combine(artifactsDir, "FormattedParagraph.docx");
        doc.Save(outputPath, SaveFormat.Docx);

        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
