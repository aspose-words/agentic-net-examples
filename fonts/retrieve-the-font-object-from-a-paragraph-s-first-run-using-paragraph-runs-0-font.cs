using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Get the first paragraph (the document always contains at least one).
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Create a run with some text and set its font name.
        Run run = new Run(doc, "Hello Aspose.Words!");
        run.Font.Name = "Arial";

        // Add the run to the paragraph.
        paragraph.AppendChild(run);

        // Retrieve the Font object from the paragraph's first run.
        Aspose.Words.Font firstRunFont = paragraph.Runs[0].Font;

        // Output a few font properties to verify the retrieval.
        Console.WriteLine("Font name: " + firstRunFont.Name);
        Console.WriteLine("Font size: " + firstRunFont.Size);
        Console.WriteLine("Bold: " + firstRunFont.Bold);
    }
}
