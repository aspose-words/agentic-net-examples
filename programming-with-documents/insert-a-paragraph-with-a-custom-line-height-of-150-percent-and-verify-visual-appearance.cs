using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsParagraphLineHeight
{
    public class Program
    {
        public static void Main()
        {
            // Define output folder and file.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);
            string outputPath = Path.Combine(outputDir, "ParagraphLineHeight.docx");

            // 1. Create a new blank document.
            Document doc = new Document();

            // 2. Use DocumentBuilder to insert a paragraph with custom line height (150% of default).
            DocumentBuilder builder = new DocumentBuilder(doc);
            // Set line spacing rule to Multiple (percentage of default line spacing).
            builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
            // Default line spacing is 12 points; 150% => 1.5 * 12 = 18 points.
            builder.ParagraphFormat.LineSpacing = 18;
            builder.Writeln("This paragraph has a line height of 150% (1.5× default).");

            // 3. Save the document.
            doc.Save(outputPath);

            // 4. Reload the document to verify that the formatting was persisted.
            Document loadedDoc = new Document(outputPath);
            Paragraph firstParagraph = loadedDoc.FirstSection.Body.FirstParagraph;

            bool lineHeightCorrect =
                firstParagraph.ParagraphFormat.LineSpacingRule == LineSpacingRule.Multiple &&
                Math.Abs(firstParagraph.ParagraphFormat.LineSpacing - 18) < 0.01;

            Console.WriteLine(lineHeightCorrect
                ? "Verification passed: line height is correctly set to 150%."
                : "Verification failed: line height is not set as expected.");
        }
    }
}
