using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the paragraph line spacing to double (2.0 lines).
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        builder.ParagraphFormat.LineSpacing = 2.0;

        // Add some text to the paragraph.
        builder.Writeln("This paragraph uses double line spacing.");

        // Save the document to a file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DoubleLineSpacing.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document saved successfully: " + outputPath);
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }

        // Validate that the paragraph line spacing is set to double.
        Paragraph firstParagraph = doc.FirstSection.Body.Paragraphs[0];
        bool isDoubleSpacing = firstParagraph.ParagraphFormat.LineSpacingRule == LineSpacingRule.Multiple &&
                               Math.Abs(firstParagraph.ParagraphFormat.LineSpacing - 2.0) < 0.0001;

        Console.WriteLine("Line spacing set to double: " + isDoubleSpacing);
    }
}
