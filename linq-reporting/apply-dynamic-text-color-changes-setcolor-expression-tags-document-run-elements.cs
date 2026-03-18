using System;
using System.Drawing;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Drawing;

class DynamicColorExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write sample text that contains custom <<setColor [Expression]>> tags.
        // The tag changes the color of all following text until another tag appears.
        builder.Writeln("Normal text. <<setColor Red>>This text should be red. ");
        builder.Writeln("More red text. <<setColor Blue>>Now blue text. ");
        builder.Writeln("Back to default color.");

        // Regular expression to match the custom color tag.
        // Example tag: <<setColor Red>>
        Regex tagRegex = new Regex(@"<<setColor\s+(\w+)>>", RegexOptions.Compiled);

        // Keep track of the current color to apply to runs.
        Color currentColor = Color.Empty; // Empty means "automatic" (no explicit color).

        // Iterate over all runs in the document.
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            string text = run.Text;
            // Find all tags inside this run.
            MatchCollection matches = tagRegex.Matches(text);
            foreach (Match match in matches)
            {
                // Extract the color name from the tag.
                string colorName = match.Groups[1].Value;
                // Try to convert the name to a System.Drawing.Color.
                // If the name is not known, fallback to Empty (automatic).
                Color parsed = Color.FromName(colorName);
                currentColor = parsed.IsKnownColor ? parsed : Color.Empty;
            }

            // Remove all tags from the run's text.
            text = tagRegex.Replace(text, string.Empty);
            run.Text = text;

            // Apply the current color to the run if a color has been set.
            if (currentColor != Color.Empty)
                run.Font.Color = currentColor;
            else
                run.Font.Color = Color.Empty; // Reset to automatic color.
        }

        // Save the resulting document.
        doc.Save("DynamicColorOutput.docx");
    }
}
