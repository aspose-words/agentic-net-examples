using Aspose.Words;
using Aspose.Words.Replacing;
using System.Text.RegularExpressions;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Create FindReplaceOptions. The FindReplaceOptions class does not have a
        // UseRegularExpressions property; using the Regex overload of Replace
        // automatically enables regular‑expression mode.
        FindReplaceOptions options = new FindReplaceOptions();

        // Optional: enable substitution handling and disable legacy mode for full feature support.
        options.UseSubstitutions = true;
        options.LegacyMode = false;

        // Define a regular‑expression pattern and its replacement text.
        Regex pattern = new Regex(@"\bcat\b", RegexOptions.IgnoreCase);
        string replacement = "dog";

        // Perform the find‑and‑replace operation using the regex overload.
        doc.Range.Replace(pattern, replacement, options);

        // Save the updated document.
        doc.Save("Output.docx");
    }
}
