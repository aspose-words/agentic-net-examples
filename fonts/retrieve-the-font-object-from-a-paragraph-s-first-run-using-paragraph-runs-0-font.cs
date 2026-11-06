using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();

        // Add a paragraph to the document.
        Paragraph paragraph = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(paragraph);

        // Add a run with sample text.
        Run run = new Run(doc, "Sample text for font retrieval.");
        paragraph.AppendChild(run);

        // Retrieve the Font object from the first run of the paragraph.
        Aspose.Words.Font font = paragraph.Runs[0].Font;

        // Optionally modify font properties.
        font.Name = "Arial";
        font.Size = 14;

        // Output font properties to the console.
        Console.WriteLine("Font Name: " + font.Name);
        Console.WriteLine("Font Size: " + font.Size);

        // Save the document to verify output.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Confirm that the file was created.
        if (System.IO.File.Exists(outputPath))
        {
            Console.WriteLine("Document saved successfully to " + outputPath);
        }
    }
}
