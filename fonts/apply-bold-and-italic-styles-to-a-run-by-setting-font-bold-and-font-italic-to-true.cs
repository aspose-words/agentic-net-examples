using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Add a new paragraph to the document body.
        Paragraph paragraph = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(paragraph);

        // Create a run with sample text.
        Run run = new Run(doc, "This text is bold and italic.");

        // Apply bold and italic styles using the Font properties.
        run.Font.Bold = true;
        run.Font.Italic = true;

        // Append the run to the paragraph.
        paragraph.AppendChild(run);

        // Validate that the styles were applied.
        bool isBold = run.Font.Bold;
        bool isItalic = run.Font.Italic;
        Console.WriteLine($"Bold applied: {isBold}, Italic applied: {isItalic}");

        // Save the document to a file.
        string outputPath = "StyledText.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        Console.WriteLine(File.Exists(outputPath) ? "Document saved successfully." : "Failed to save document.");
    }
}
