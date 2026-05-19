using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Aspose.Words.Document doc = new Aspose.Words.Document();

        // Use DocumentBuilder to add a paragraph with a single run of text.
        Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
        builder.Writeln("Hello World!");

        // Retrieve the first paragraph in the document.
        Aspose.Words.Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Retrieve the Font object from the paragraph's first run.
        Aspose.Words.Font firstRunFont = paragraph.Runs[0].Font;

        // Output some font properties to verify retrieval.
        Console.WriteLine("First run font name: " + firstRunFont.Name);
        Console.WriteLine("First run font size: " + firstRunFont.Size);

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);
    }
}
