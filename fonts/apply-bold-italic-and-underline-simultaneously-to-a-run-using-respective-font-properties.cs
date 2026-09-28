using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with a single run of text.
        builder.Write("Sample text for font styling.");

        // Retrieve the run we just added.
        Paragraph paragraph = doc.FirstSection.Body.Paragraphs[0];
        Run run = (Run)paragraph.Runs[0];

        // Apply bold, italic, and underline formatting.
        run.Font.Bold = true;
        run.Font.Italic = true;
        // The correct enum for underline style is Underline (not UnderlineType).
        run.Font.Underline = Underline.Single;

        // Validate that the properties were set correctly.
        bool isBold = run.Font.Bold;
        bool isItalic = run.Font.Italic;
        bool isUnderline = run.Font.Underline == Underline.Single;

        if (isBold && isItalic && isUnderline)
        {
            Console.WriteLine("Bold, italic, and underline applied successfully.");
        }
        else
        {
            Console.WriteLine("Font formatting validation failed.");
        }

        // Save the document to disk.
        string outputPath = "FormattedRun.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully at: {Path.GetFullPath(outputPath)}");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
