using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a paragraph and add it to the document's first section body.
        Paragraph paragraph = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(paragraph);

        // Create a run with some text and set its font name.
        Run run = new Run(doc, "Hello Aspose.Words!");
        run.Font.Name = "Courier New";
        paragraph.AppendChild(run);

        // Retrieve the Font object from the first run of the paragraph.
        Font firstRunFont = paragraph.Runs[0].Font;

        // Output a property of the retrieved font.
        Console.WriteLine("First run font name: " + firstRunFont.Name);

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FirstRunFont.docx");
        doc.Save(outputPath);
    }
}
