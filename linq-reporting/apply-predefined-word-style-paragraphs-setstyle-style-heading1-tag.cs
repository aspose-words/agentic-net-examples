using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyStyleFromTag
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Regular expression to capture the style name inside the tag.
        // Example tag: <<setStyle style="Heading1">>
        Regex tagRegex = new Regex(@"<<setStyle\s+style\s*=\s*""(?<style>[^""]+)""\s*>>", RegexOptions.Compiled);

        // Iterate over all paragraphs in the document.
        foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // Search each run within the paragraph for the custom tag.
            foreach (Run run in paragraph.Runs)
            {
                if (run.Text == null) continue;

                Match match = tagRegex.Match(run.Text);
                if (match.Success)
                {
                    // Extract the style name from the tag.
                    string styleName = match.Groups["style"].Value;

                    // Apply the extracted style to the whole paragraph.
                    paragraph.ParagraphFormat.StyleName = styleName;

                    // Remove the tag text from the run.
                    run.Text = tagRegex.Replace(run.Text, string.Empty);
                }
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
