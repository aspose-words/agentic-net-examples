using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class FindReplaceInFields
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Configure find/replace options to include fields in the search.
        FindReplaceOptions options = new FindReplaceOptions();
        // By default fields are included, but we set it explicitly for clarity.
        options.IgnoreFields = false;          // Search inside field codes and results.
        options.IgnoreFieldCodes = false;      // Include the field code text itself.
        options.MatchCase = false;            // Case‑insensitive search (adjust as needed).

        // Define the text to find and its replacement.
        string oldText = "PLACEHOLDER";
        string newText = "ActualValue";

        // Perform the replace operation across the whole document range.
        doc.Range.Replace(oldText, newText, options);

        // Optionally update all fields so that any calculated results reflect the new data.
        doc.UpdateFields();

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
