using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsParagraphStyleExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Attach a DocumentBuilder to the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set the left indent to 0.5 inches (0.5 * 72 points per inch = 36 points).
            builder.ParagraphFormat.LeftIndent = 36.0;

            // Set line spacing to 1.5 lines.
            builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
            builder.ParagraphFormat.LineSpacing = 1.5;

            // Write a sample paragraph that will use the custom formatting.
            builder.Writeln("This paragraph has a left indent of 0.5 inches and line spacing of 1.5.");

            // Save the document to the local file system.
            string outputPath = "CustomParagraphStyle.docx";
            doc.Save(outputPath);
        }
    }
}
