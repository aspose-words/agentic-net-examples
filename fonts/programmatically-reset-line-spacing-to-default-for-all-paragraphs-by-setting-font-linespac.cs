using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fonts;

namespace ResetLineSpacingExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to add sample paragraphs with custom line spacing.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // First paragraph – set a larger line spacing.
            builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.AtLeast;
            builder.ParagraphFormat.LineSpacing = 30; // 30 points.
            builder.Writeln("This paragraph has a custom line spacing of 30 points.");

            // Second paragraph – set a different line spacing.
            builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Exactly;
            builder.ParagraphFormat.LineSpacing = 10; // 10 points.
            builder.Writeln("This paragraph has an exact line spacing of 10 points.");

            // Reset the paragraph formatting to default for all paragraphs in the document.
            foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                // Reset line spacing to the default value (0) and use the default rule (Multiple).
                paragraph.ParagraphFormat.LineSpacing = 0;
                paragraph.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
            }

            // Define the output file path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ResetLineSpacing.docx");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (File.Exists(outputPath))
            {
                Console.WriteLine($"Document saved successfully to: {outputPath}");
            }
            else
            {
                Console.WriteLine("Failed to save the document.");
            }
        }
    }
}
