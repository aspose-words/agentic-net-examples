using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing; // Required package, not used directly
using Newtonsoft.Json; // Required package, not used directly

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Aspose.Words.Document doc = new Aspose.Words.Document();

        // Initialize a DocumentBuilder to add content to the document.
        Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);

        // Add a paragraph of text.
        builder.Writeln("This paragraph will have double line spacing.");

        // Configure the paragraph to use multiple line spacing (based on line count).
        builder.ParagraphFormat.LineSpacingRule = Aspose.Words.LineSpacingRule.Multiple;

        // Set the line spacing to double (2 * 12 points = 24 points).
        builder.ParagraphFormat.LineSpacing = 24.0;

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DoubleLineSpacing.docx");
        doc.Save(outputPath);

        // Verify that the file was created successfully.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document saved successfully: " + outputPath);
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }

        // Output the line spacing value to confirm it was set.
        double lineSpacing = builder.ParagraphFormat.LineSpacing;
        Console.WriteLine("Line spacing set to: " + lineSpacing);
    }
}
