using System;
using Aspose.Words;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Create a new document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply an emphasis mark to the font of the upcoming run.
        // For example, use OverSolidCircle (value = 1).
        builder.Font.EmphasisMark = EmphasisMark.OverSolidCircle;

        // Write a run of text that will carry the emphasis mark.
        builder.Writeln("Text with emphasis mark");

        // Save the document to a temporary file.
        string filePath = "EmphasisMarkDemo.docx";
        doc.Save(filePath);

        // Load the document back from the file.
        Document loadedDoc = new Document(filePath);

        // Retrieve the first run in the document.
        Run firstRun = (Run)loadedDoc.FirstSection.Body.FirstParagraph.GetChild(NodeType.Run, 0, true);

        // Get the EmphasisMark value from the run's font.
        EmphasisMark mark = firstRun.Font.EmphasisMark;

        // Display the retrieved value for debugging.
        Console.WriteLine($"EmphasisMark of the first run: {mark} (numeric value = {(int)mark})");
    }
}
