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
        // Create a sample document with email addresses.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Please contact us at support@example.com for assistance.");
        builder.Writeln("You can also reach sales@example.org or info@example.net.");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loadedDoc = new Document(inputPath);

        // Define email regex pattern.
        Regex emailRegex = new Regex(@"\b[\w\.-]+@[\w\.-]+\.\w{2,}\b", RegexOptions.Compiled);

        // Set up the replacing callback that only records matches.
        EmailCollectorCallback collector = new EmailCollectorCallback();
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = collector,
            MatchCase = false
        };

        // Perform the replacement (no actual text change, just counting matches).
        int replacedCount = loadedDoc.Range.Replace(emailRegex, "$0", options);
        if (replacedCount == 0)
        {
            throw new InvalidOperationException("No email addresses were found for replacement.");
        }

        // Insert hyperlinks for each recorded email address.
        InsertHyperlinks(loadedDoc, collector.ReplacedEmails);

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);

        // Write a JSON report of replaced emails.
        string jsonReport = JsonConvert.SerializeObject(collector.ReplacedEmails, Formatting.Indented);
        File.WriteAllText("replacements.json", jsonReport);
    }

    private static void InsertHyperlinks(Document doc, List<string> emails)
    {
        // Build a hash set for quick lookup.
        HashSet<string> emailSet = new HashSet<string>(emails);

        // Collect runs that exactly match an email address.
        List<Run> runsToReplace = new List<Run>();
        NodeCollection runs = doc.GetChildNodes(NodeType.Run, true);
        foreach (Run run in runs)
        {
            if (emailSet.Contains(run.Text))
            {
                runsToReplace.Add(run);
            }
        }

        // Replace each run with a hyperlink.
        foreach (Run run in runsToReplace)
        {
            string email = run.Text;
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveTo(run);
            builder.InsertHyperlink(email, "mailto:" + email, false);
            run.Remove();
        }
    }

    private class EmailCollectorCallback : IReplacingCallback
    {
        public List<string> ReplacedEmails { get; } = new List<string>();

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Record the matched email address.
            string email = args.Match.Value;
            ReplacedEmails.Add(email);

            // Keep the original text unchanged.
            args.Replacement = email;
            return ReplaceAction.Replace;
        }
    }
}
