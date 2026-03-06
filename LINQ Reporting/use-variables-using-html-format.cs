using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph that will contain an HTML fragment with a placeholder.
        builder.Writeln("HTML with variable placeholder:");
        string htmlFragment = "<p>Variable content: ${Title}</p>";
        builder.InsertHtml(htmlFragment);

        // Add a document variable whose value is HTML formatted text.
        doc.Variables.Add("Title", "<b>Bold Text</b>");

        // Replace the placeholder with the variable value, interpreting the replacement as HTML.
        FindReplaceOptions replaceOptions = new FindReplaceOptions
        {
            ReplacementFormat = ReplacementFormat.Html
        };
        doc.Range.Replace("${Title}", doc.Variables["Title"], replaceOptions);

        // Save the resulting document.
        doc.Save("VariablesWithHtml.docx");
    }
}
