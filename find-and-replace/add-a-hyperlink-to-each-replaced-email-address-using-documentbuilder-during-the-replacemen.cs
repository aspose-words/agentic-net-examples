using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing email addresses.
        const string inputPath = "input.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(
            "Please contact us at support@example.com or sales@example.com for more information.");
        doc.Save(inputPath);

        // Load the document for processing.
        Document loadedDoc = new Document(inputPath);

        // Define a regular expression that matches e‑mail addresses.
        Regex emailRegex = new Regex(@"\b[\w\.-]+@[\w\.-]+\.\w{2,}\b", RegexOptions.Compiled);

        // Set up a callback that replaces each e‑mail with a hyperlink.
        EmailHyperlinkCallback callback = new EmailHyperlinkCallback();

        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = callback
        };

        // Perform the replacement.
        int replacedCount = loadedDoc.Range.Replace(emailRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (callback.ReplacedEmails.Count == 0)
            throw new InvalidOperationException("No email addresses were replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);

        // Write a JSON report of the replaced e‑mail addresses.
        const string reportPath = "report.json";
        string json = JsonConvert.SerializeObject(callback.ReplacedEmails, Formatting.Indented);
        File.WriteAllText(reportPath, json);
    }

    private class EmailHyperlinkCallback : IReplacingCallback
    {
        public List<string> ReplacedEmails { get; } = new List<string>();

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Get the matched e‑mail address.
            string email = args.Match.Value;
            ReplacedEmails.Add(email);

            // The node that contains the match is a Run.
            // Insert a hyperlink into the same paragraph and then remove the original run.
            Document doc = (Document)args.MatchNode.Document;
            Run runNode = args.MatchNode as Run;
            if (runNode == null)
                return ReplaceAction.Skip; // Safety check.

            Paragraph paragraph = runNode.ParentNode as Paragraph;
            if (paragraph == null)
                return ReplaceAction.Skip; // Safety check.

            // Insert the hyperlink just before the original run.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveTo(paragraph);
            // Move the cursor to the position of the original run.
            builder.MoveTo(runNode);
            builder.InsertHyperlink(email, "mailto:" + email, false);

            // Remove the original plain e‑mail text run.
            runNode.Remove();

            // Skip the default replacement because we already inserted the hyperlink.
            return ReplaceAction.Skip;
        }
    }
}
