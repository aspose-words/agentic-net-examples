using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Add a new paragraph to the document.
        Paragraph paragraph = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(paragraph);

        // Add a run with some text to the paragraph.
        Run run = new Run(doc, "Hello Aspose.Words!");
        // Set some font properties for demonstration.
        run.Font.Name = "Arial";
        run.Font.Size = 16;
        paragraph.AppendChild(run);

        // Retrieve the Font object from the first run of the paragraph.
        Aspose.Words.Font firstRunFont = paragraph.Runs[0].Font;

        // Output font properties to the console.
        Console.WriteLine("Font Name: " + firstRunFont.Name);
        Console.WriteLine("Font Size: " + firstRunFont.Size);

        // Save the document to verify that an output file is created.
        string outputPath = "OutputDocument.docx";
        doc.Save(outputPath);

        // Confirm that the file was saved.
        if (System.IO.File.Exists(outputPath))
        {
            Console.WriteLine("Document saved successfully: " + outputPath);
        }
    }
}
