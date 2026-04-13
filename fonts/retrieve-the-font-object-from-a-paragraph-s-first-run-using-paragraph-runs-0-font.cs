using System;
using Aspose.Words;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a paragraph to the document body.
        Paragraph paragraph = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(paragraph);

        // Create a run with some text.
        Run run = new Run(doc, "Hello Aspose!");
        paragraph.AppendChild(run);

        // Set font properties for the run.
        Aspose.Words.Font runFont = run.Font;
        runFont.Name = "Arial";
        runFont.Size = 24;

        // Use Aspose.Drawing.Color and convert to System.Drawing.Color for the Font.Color property.
        Aspose.Drawing.Color aspColor = Aspose.Drawing.Color.Red;
        runFont.Color = System.Drawing.Color.FromArgb(aspColor.ToArgb());

        // Retrieve the Font object from the paragraph's first run.
        Aspose.Words.Font retrievedFont = paragraph.Runs[0].Font;

        // Output some font properties to the console.
        Console.WriteLine($"Font Name: {retrievedFont.Name}");
        Console.WriteLine($"Font Size: {retrievedFont.Size}");
        Console.WriteLine($"Font Color ARGB: {retrievedFont.Color.ToArgb()}");

        // Save the document to a file.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
