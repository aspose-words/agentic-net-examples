using System;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Add variables that contain Markdown formatted text.
        doc.Variables.Add("Title", "## Sample Title");
        doc.Variables.Add("Author", "*John Doe*");

        // Insert placeholders where the variables will be substituted.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("{{Title}}");
        builder.Writeln("Written by {{Author}}");

        // Configure replace options to treat the replacement string as Markdown.
        FindReplaceOptions replaceOptions = new FindReplaceOptions
        {
            ReplacementFormat = ReplacementFormat.Markdown
        };

        // Replace placeholders with the corresponding variable values.
        doc.Range.Replace("{{Title}}", doc.Variables["Title"], replaceOptions);
        doc.Range.Replace("{{Author}}", doc.Variables["Author"], replaceOptions);

        // Save the resulting document as a Markdown file.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();
        doc.Save("Output.md", saveOptions);
    }
}
