using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace ResetLineSpacingExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Get the first (and only) section and its body.
            Section section = doc.Sections[0];
            Body body = section.Body;

            // Paragraph with custom line spacing (30 points, exactly).
            Paragraph para1 = new Paragraph(doc);
            Run run1 = new Run(doc, "This paragraph has custom line spacing.");
            para1.ParagraphFormat.LineSpacing = 30.0;
            para1.ParagraphFormat.LineSpacingRule = Aspose.Words.LineSpacingRule.Exactly;
            para1.AppendChild(run1);
            body.AppendChild(para1);

            // Paragraph that uses the default line spacing.
            Paragraph para2 = new Paragraph(doc);
            Run run2 = new Run(doc, "This paragraph uses the default line spacing.");
            para2.AppendChild(run2);
            body.AppendChild(para2);

            // Reset line spacing for all paragraphs to the default behavior.
            // Setting LineSpacing to 0 restores the default; the rule does not need to be set explicitly.
            foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                paragraph.ParagraphFormat.LineSpacing = 0;
                // No need to set LineSpacingRule because a value of 0 makes the rule irrelevant.
            }

            // Validate that every paragraph now has the default line‑spacing setting (LineSpacing == 0).
            foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                if (paragraph.ParagraphFormat.LineSpacing != 0)
                {
                    throw new InvalidOperationException("Line spacing reset validation failed.");
                }
            }

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
            doc.Save(outputPath);

            // Ensure the file was created.
            if (!File.Exists(outputPath))
            {
                throw new FileNotFoundException("The output document was not created.", outputPath);
            }

            // Indicate success.
            Console.WriteLine("Document saved successfully to: " + outputPath);
        }
    }
}
