using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a new paragraph to the document.
        Paragraph paragraph = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(paragraph);

        // Create a run with sample text.
        Run run = new Run(doc, "This text is underlined.");
        paragraph.AppendChild(run);

        // Apply single underline style to the run.
        run.Font.Underline = Aspose.Words.Underline.Single;

        // Validate that the underline style was applied.
        if (run.Font.Underline == Aspose.Words.Underline.Single)
        {
            Console.WriteLine("Underline applied successfully.");
        }
        else
        {
            Console.WriteLine("Failed to apply underline.");
        }

        // Save the document to a file.
        string outputPath = "UnderlineExample.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved to {outputPath}");
        }
        else
        {
            Console.WriteLine("Document was not saved.");
        }
    }
}
