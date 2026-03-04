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

        // Add a document variable whose value contains HTML markup.
        doc.Variables.Add("MyVar", "Hello <b>World</b>");

        // Insert an HTML fragment that includes a placeholder for the variable.
        string html = "<p>Variable content: {{MyVar}}</p>";
        builder.InsertHtml(html);

        // Prepare find‑replace options to treat the replacement string as HTML.
        FindReplaceOptions replaceOptions = new FindReplaceOptions();
        replaceOptions.ReplacementFormat = ReplacementFormat.Html;

        // Replace the placeholder with the variable's value, preserving its HTML formatting.
        string variableValue = doc.Variables["MyVar"];
        doc.Range.Replace("{{MyVar}}", variableValue, replaceOptions);

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
