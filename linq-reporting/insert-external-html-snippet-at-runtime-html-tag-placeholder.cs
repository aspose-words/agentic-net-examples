using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load a template document that contains the placeholder <<html>>.
        Document doc = new Document("Template.docx");

        // HTML snippet that will be inserted at runtime.
        const string htmlSnippet = "<h2>Dynamic Content</h2><p>This content was inserted programmatically.</p>";

        // Configure find‑replace options with a custom callback.
        FindReplaceOptions replaceOptions = new FindReplaceOptions();
        replaceOptions.ReplacingCallback = new HtmlInsertCallback(htmlSnippet);

        // Replace the placeholder. The callback will insert the HTML at the match location.
        doc.Range.Replace(new Regex(@"<<html>>"), string.Empty, replaceOptions);

        // Save the modified document.
        doc.Save("Result.docx");
    }

    // Callback that inserts the HTML snippet where the placeholder is found.
    private class HtmlInsertCallback : IReplacingCallback
    {
        private readonly string _html;

        public HtmlInsertCallback(string html) => _html = html;

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Position the builder at the start of the matched placeholder node.
            // In newer Aspose.Words versions args.MatchNode.Document returns DocumentBase.
            // Cast it to Document, which is the type expected by DocumentBuilder's constructor.
            DocumentBuilder builder = new DocumentBuilder((Document)args.MatchNode.Document);
            builder.MoveTo(args.MatchNode);

            // Insert the HTML fragment.
            builder.InsertHtml(_html);

            // Skip the default text replacement because we have already handled the insertion.
            return ReplaceAction.Skip;
        }
    }
}
