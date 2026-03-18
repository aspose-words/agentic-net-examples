using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Template.docx");

        // Regex that matches the custom tag, e.g. <setAlignment alignment="center">
        Regex tagRegex = new Regex(@"<setAlignment\s+alignment\s*=\s*""(?<align>\w+)""\s*>",
                                   RegexOptions.IgnoreCase);

        // Iterate over every paragraph in the document.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            string paragraphText = para.GetText();

            // Look for the alignment tag inside the paragraph.
            Match match = tagRegex.Match(paragraphText);
            if (!match.Success) continue;

            // Extract the alignment value (e.g., "center") and convert it to the enum.
            string alignValue = match.Groups["align"].Value;
            ParagraphAlignment alignment = ParagraphAlignment.Left; // fallback

            if (Enum.TryParse<ParagraphAlignment>(alignValue, true, out ParagraphAlignment parsed))
                alignment = parsed;

            // Apply the alignment to the current paragraph.
            para.ParagraphFormat.Alignment = alignment;

            // Remove the tag from the paragraph text.
            string cleanedText = tagRegex.Replace(paragraphText, string.Empty);

            // Replace the paragraph's runs with a single run containing the cleaned text.
            para.RemoveAllChildren();
            para.AppendChild(new Run(doc, cleanedText));
        }

        // Save the updated document.
        doc.Save("Result.docx");
    }
}
