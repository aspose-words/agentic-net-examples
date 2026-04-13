using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a new run with some text.
        Run run = new Run(doc, "Hello Aspose.Words!");

        // Set the font name of the run to Arial using the Font.Name property.
        Aspose.Words.Font font = run.Font;
        font.Name = "Arial";

        // Validate that the font name was set correctly.
        if (font.Name != "Arial")
        {
            throw new InvalidOperationException("Font name was not set to Arial.");
        }

        // Ensure the document has at least one paragraph to host the run.
        if (doc.FirstSection?.Body?.FirstParagraph == null)
        {
            Paragraph para = new Paragraph(doc);
            doc.FirstSection.Body.AppendChild(para);
        }

        // Append the run to the first paragraph.
        doc.FirstSection.Body.FirstParagraph.AppendChild(run);

        // Define the output file path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "FormattedRun.docx");

        // Save the document to the file system.
        doc.Save(outputPath);

        // Verify that the output file exists.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The document was not saved correctly.", outputPath);
        }

        // Inform that the process completed successfully.
        Console.WriteLine($"Document saved successfully to: {outputPath}");
    }
}
