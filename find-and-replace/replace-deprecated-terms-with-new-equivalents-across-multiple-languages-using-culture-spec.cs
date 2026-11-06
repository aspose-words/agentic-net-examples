using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class ReplacementRecord
{
    public string Language { get; set; } = string.Empty;
    public string OldValue { get; set; } = string.Empty;
    public string NewValue { get; set; } = string.Empty;
}

public class ReplacementLogger : IReplacingCallback
{
    public List<ReplacementRecord> Records { get; } = new List<ReplacementRecord>();
    public string CurrentLanguage { get; set; } = string.Empty;
    public string NewValue { get; set; } = string.Empty;

    ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
    {
        Records.Add(new ReplacementRecord
        {
            Language = CurrentLanguage,
            OldValue = args.Match.Value,
            NewValue = NewValue
        });
        return ReplaceAction.Replace;
    }
}

public class Program
{
    public static void Main()
    {
        // Create a sample document with deprecated terms in different languages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The old term foo is deprecated.");
        builder.Writeln("Le terme ancien est obsolète."); // French: "ancien"
        builder.Writeln("Der alte Begriff veraltet ist veraltet."); // German: "veraltet"
        builder.Writeln("El término viejo foo está obsoleto."); // Spanish: "foo"
        doc.Save("input.docx");

        // Load the document for processing.
        Document loadedDoc = new Document("input.docx");

        // Define language‑specific patterns and replacements.
        var replacements = new List<(string Language, Regex Pattern, string NewTerm)>
        {
            ("English", new Regex(@"\bfoo\b", RegexOptions.IgnoreCase), "bar"),
            ("French", new Regex(@"\ban(ancien|cien)\b", RegexOptions.IgnoreCase), "nouveau"),
            ("German", new Regex(@"\bveraltet\b", RegexOptions.IgnoreCase), "aktuell"),
            ("Spanish", new Regex(@"\bfoo\b", RegexOptions.IgnoreCase), "bar")
        };

        var logger = new ReplacementLogger();
        int totalReplacements = 0;

        foreach (var (language, pattern, newTerm) in replacements)
        {
            logger.CurrentLanguage = language;
            logger.NewValue = newTerm;

            FindReplaceOptions options = new FindReplaceOptions
            {
                ReplacingCallback = logger,
                MatchCase = false,
                FindWholeWordsOnly = true
            };

            int replaced = loadedDoc.Range.Replace(pattern, newTerm, options);
            totalReplacements += replaced;
        }

        if (totalReplacements == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        loadedDoc.Save("output.docx");

        // Serialize replacement log to JSON.
        string jsonReport = JsonConvert.SerializeObject(logger.Records, Formatting.Indented);
        File.WriteAllText("report.json", jsonReport);

        // Validate that the report file was created.
        if (!File.Exists("report.json"))
            throw new InvalidOperationException("Report file was not created.");
    }
}
