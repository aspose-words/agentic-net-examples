using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareDocumentsExample
{
    static void Main()
    {
        // Folder where the resulting document will be saved.
        string artifactsDir = "Artifacts/";
        // Ensure the folder exists.
        System.IO.Directory.CreateDirectory(artifactsDir);

        // ---------- Create the original document ----------
        Document docOriginal = new Document();
        DocumentBuilder builder = new DocumentBuilder(docOriginal);
        builder.Writeln("Alpha Lorem ipsum dolor sit amet, consectetur adipiscing elit.");
        builder.Writeln("Second paragraph with some text.");

        // ---------- Create the edited document ----------
        Document docEdited = new Document();
        DocumentBuilder builderEdited = new DocumentBuilder(docEdited);
        // Slightly modify the first line to demonstrate character‑level differences.
        builderEdited.Writeln("AlphA Lorem ipsum dolor sit amet, consectetur adipiscing elit!");
        builderEdited.Writeln("Second paragraph with some changed text.");

        // ---------- Configure comparison options ----------
        CompareOptions compareOptions = new CompareOptions
        {
            // Track changes at the character level.
            Granularity = Granularity.CharLevel,
            // Use the edited document as the target (equivalent to Word's "Show changes in New").
            Target = ComparisonTargetType.New,
            // Keep other options at their defaults (no ignoring).
            CompareMoves = false,
            IgnoreFormatting = false,
            IgnoreCaseChanges = false,
            IgnoreComments = false,
            IgnoreTables = false,
            IgnoreFields = false,
            IgnoreFootnotes = false,
            IgnoreTextboxes = false,
            IgnoreHeadersAndFooters = false
        };

        // ---------- Perform the comparison ----------
        // The revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);

        // ---------- Save the result ----------
        docOriginal.Save(artifactsDir + "ComparisonResult.docx");
    }
}
