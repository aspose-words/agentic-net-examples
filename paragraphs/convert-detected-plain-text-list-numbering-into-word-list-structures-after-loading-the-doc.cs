using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and add sample paragraphs with plain‑text numbering.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("1. First item of a list");
        builder.Writeln("2. Second item of a list");
        builder.Writeln("3. Third item of a list");
        builder.Writeln("This is a normal paragraph without numbering.");
        builder.Writeln("4. Fourth item after a normal paragraph");
        builder.Writeln("5. Fifth item");

        // Create a list that will be used for the converted items.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Regular expression to detect plain‑text numbering at the start of a paragraph.
        Regex numberingRegex = new Regex(@"^\d+\.\s", RegexOptions.Compiled);

        // Traverse all paragraphs in the document.
        foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // Get the paragraph text (includes the paragraph mark \r at the end).
            string paragraphText = paragraph.GetText();

            // Trim leading whitespace for detection.
            string trimmedStart = paragraphText.TrimStart();

            // If the paragraph starts with a number followed by a dot and a space, convert it.
            if (numberingRegex.IsMatch(trimmedStart))
            {
                // Remove the plain‑text number prefix.
                string cleanedText = numberingRegex.Replace(paragraphText, string.Empty);

                // Update the paragraph's first run (or create one) with the cleaned text,
                // ensuring the paragraph mark (\r) is not part of the run text.
                string newRunText = cleanedText.TrimEnd('\r');

                Run firstRun = paragraph.FirstChild as Run;
                if (firstRun != null)
                {
                    firstRun.Text = newRunText;
                }
                else
                {
                    paragraph.AppendChild(new Run(doc, newRunText));
                }

                // Apply the list formatting to this paragraph.
                paragraph.ListFormat.List = list;
                paragraph.ListFormat.ListLevelNumber = 0;
            }
        }

        // Save the resulting document.
        doc.Save("ConvertedList.docx");
    }
}
