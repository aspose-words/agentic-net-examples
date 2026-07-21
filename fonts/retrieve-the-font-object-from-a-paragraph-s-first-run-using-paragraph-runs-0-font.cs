using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a new paragraph and add it to the document's body.
        Paragraph paragraph = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(paragraph);

        // Create a run with some text and set its font properties.
        Run run = new Run(doc, "Hello Aspose!");
        run.Font.Name = "Arial";
        run.Font.Size = 24;
        paragraph.AppendChild(run);

        // Retrieve the Font object from the paragraph's first run.
        Aspose.Words.Font firstRunFont = paragraph.Runs[0].Font;

        // Output some font information to the console.
        Console.WriteLine($"First run font name: {firstRunFont.Name}");
        Console.WriteLine($"First run font size: {firstRunFont.Size}");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FirstRunFont.docx");
        doc.Save(outputPath);
    }
}
