using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Insert a paragraph with a run of text.
        Paragraph para = doc.FirstSection.Body.FirstParagraph;
        Run run = new Run(doc, "Hello world!");
        para.AppendChild(run);

        // Access the font of the run and retrieve its size (default is 12 points).
        double fontSize = run.Font.Size;

        // Output the font size for debugging purposes.
        Console.WriteLine($"Run font size: {fontSize} pt");
    }
}
