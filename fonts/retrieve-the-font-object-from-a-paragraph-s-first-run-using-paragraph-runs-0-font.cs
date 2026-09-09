using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a paragraph and a run with some text.
        Paragraph paragraph = new Paragraph(doc);
        Run run = new Run(doc, "Hello Aspose!");
        run.Font.Name = "Arial";
        run.Font.Size = 24;
        paragraph.AppendChild(run);

        // Add the paragraph to the document body.
        doc.FirstSection.Body.AppendChild(paragraph);

        // Retrieve the Font object from the paragraph's first run.
        Aspose.Words.Font firstRunFont = paragraph.Runs[0].Font;

        // Display font properties to verify the retrieval.
        Console.WriteLine($"Font Name: {firstRunFont.Name}");
        Console.WriteLine($"Font Size: {firstRunFont.Size}");

        // Save the document.
        doc.Save("Output.docx");
    }
}
