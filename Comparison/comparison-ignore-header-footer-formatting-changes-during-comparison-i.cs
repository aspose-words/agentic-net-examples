using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

class HeaderFooterComparison
{
    static void Main()
    {
        // Directory for output files.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // ---------- Create the original document ----------
        Document docOriginal = new Document();
        DocumentBuilder builder = new DocumentBuilder(docOriginal);

        // Body content.
        builder.Writeln("Original body paragraph.");

        // Header content.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Original header text.");

        // ---------- Clone and modify the header ----------
        Document docEdited = (Document)docOriginal.Clone(true);
        builder = new DocumentBuilder(docEdited);

        // Change header text and formatting (these changes will be ignored).
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Font.Size = 20;               // Change font size.
        builder.Writeln("Edited header text with larger font.");

        // ---------- Set comparison options to ignore headers/footers ----------
        CompareOptions compareOptions = new CompareOptions
        {
            // Ignore any differences in header/footer content.
            IgnoreHeadersAndFooters = true,
            // Use the edited document as the base for comparison (optional).
            Target = ComparisonTargetType.New
        };

        // Perform the comparison.
        docOriginal.Compare(docEdited, "Author", DateTime.Now, compareOptions);

        // Save the result. Revisions will reflect body changes only; header changes are ignored.
        docOriginal.Save(Path.Combine(artifactsDir, "ComparisonResult.docx"));
    }
}
