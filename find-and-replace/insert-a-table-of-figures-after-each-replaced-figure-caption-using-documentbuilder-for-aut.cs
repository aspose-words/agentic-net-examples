using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document with figure captions.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Introduction paragraph.");
        builder.Writeln("Figure 1: Old caption for the first figure.");
        builder.Writeln("Some intermediate text.");
        builder.Writeln("Figure 2: Old caption for the second figure.");
        builder.Writeln("Conclusion paragraph.");

        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loadedDoc = new Document(inputPath);

        // Define the regex to find figure captions.
        Regex figureCaptionRegex = new Regex(@"Figure (\d+): .*", RegexOptions.Compiled);

        // Set up the replace callback.
        var callback = new FigureCaptionReplaceCallback(loadedDoc);
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = callback
        };

        // Perform the replace operation.
        int replacedCount = loadedDoc.Range.Replace(figureCaptionRegex, string.Empty, options);

        // Validate that replacements occurred.
        if (replacedCount == 0 || callback.ReplacementCount == 0)
            throw new InvalidOperationException("No figure captions were replaced.");

        // Update fields (e.g., the inserted Table of Figures) before saving.
        loadedDoc.UpdateFields();

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);

        // Write a simple JSON report of the replacements.
        var report = new
        {
            Replacements = callback.ReplacementCount,
            OutputDocument = Path.GetFullPath(outputPath)
        };
        string jsonReport = JsonConvert.SerializeObject(report, Formatting.Indented);
        File.WriteAllText("report.json", jsonReport);
    }
}

// Callback that replaces figure captions and inserts a Table of Figures after each replaced caption.
public class FigureCaptionReplaceCallback : IReplacingCallback
{
    private readonly Document _document;
    public int ReplacementCount { get; private set; }

    public FigureCaptionReplaceCallback(Document document)
    {
        _document = document ?? throw new ArgumentNullException(nameof(document));
    }

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Ensure we have a match.
        if (args.Match == null || !args.Match.Success)
            return ReplaceAction.Skip;

        // Extract the figure number from the match.
        string figureNumber = args.Match.Groups[1].Value;

        // Set the new caption text.
        args.Replacement = $"Figure {figureNumber}: Updated caption";

        // Locate the paragraph that contains the match.
        Node? node = args.MatchNode;
        while (node != null && !(node is Paragraph))
        {
            node = node.ParentNode;
        }

        if (node is Paragraph paragraph)
        {
            // Move the builder to the paragraph, insert a new paragraph after it,
            // and add a Table of Figures field.
            DocumentBuilder builder = new DocumentBuilder(_document);
            builder.MoveTo(paragraph);
            builder.InsertParagraph(); // Inserts a new paragraph after the current one.
            // Insert a Table of Figures field (TOC field with caption label "Figure").
            builder.InsertField(@"TOC \h ""Figure"" \c ""Figure""");
        }

        ReplacementCount++;
        return ReplaceAction.Replace;
    }
}
