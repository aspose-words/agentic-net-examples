using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to add a paragraph with a single run.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for font styling.");

        // Retrieve the first run from the paragraph that contains the text.
        // After Writeln the builder is positioned on a new empty paragraph,
        // so we need to access the previous paragraph.
        Paragraph paragraphWithRun = (Paragraph)builder.CurrentParagraph.PreviousSibling;
        Run run = paragraphWithRun.Runs[0];

        // Apply bold and italic styles.
        run.Font.Bold = true;
        run.Font.Italic = true;

        // Validate that the styles were applied.
        if (!run.Font.Bold || !run.Font.Italic)
        {
            throw new InvalidOperationException("Failed to set font styles.");
        }

        // Save the document to disk.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Ensure the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The output file was not created.", outputPath);
        }

        // Indicate success.
        Console.WriteLine("Document created successfully at " + Path.GetFullPath(outputPath));
    }
}
