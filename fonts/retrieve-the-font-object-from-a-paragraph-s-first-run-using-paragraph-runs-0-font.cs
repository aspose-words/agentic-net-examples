using System;
using System.IO;
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

        // Create a run with some text.
        Run run = new Run(doc, "Hello Aspose.Words!");

        // Set font properties for the run.
        Aspose.Words.Font runFont = run.Font;
        runFont.Name = "Arial";
        runFont.Size = 16;

        // Add the run to the paragraph.
        paragraph.AppendChild(run);

        // Save the document to a file.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Retrieve the Font object from the paragraph's first run.
        if (paragraph.Runs.Count > 0)
        {
            Aspose.Words.Font firstRunFont = paragraph.Runs[0].Font;
            // Output font properties to the console.
            Console.WriteLine("Font Name: " + firstRunFont.Name);
            Console.WriteLine("Font Size: " + firstRunFont.Size);
        }

        // Verify that the output file exists.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document saved successfully.");
        }
        else
        {
            Console.WriteLine("Failed to save document.");
        }
    }
}
