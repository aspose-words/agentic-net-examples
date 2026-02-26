using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the DOCX document that contains Structured Document Tags (SDTs).
        Document doc = new Document("StructuredDocumentTags.docx");

        // Text to search for inside the document.
        string findText = "Placeholder";
        // Replacement text.
        string replaceText = "Replaced";

        // -----------------------------------------------------------------
        // 1. Replace text while treating each SDT as a separate story.
        //    The replacement will NOT cross SDT boundaries.
        // -----------------------------------------------------------------
        FindReplaceOptions options = new FindReplaceOptions
        {
            // Do NOT ignore the content of SDTs.
            IgnoreStructuredDocumentTags = false
        };
        doc.Range.Replace(findText, replaceText, options);

        // -----------------------------------------------------------------
        // 2. Replace text while ignoring SDT boundaries.
        //    The content of each SDT is treated as simple text.
        // -----------------------------------------------------------------
        FindReplaceOptions optionsIgnore = new FindReplaceOptions
        {
            // Ignore the content of SDTs.
            IgnoreStructuredDocumentTags = true
        };
        doc.Range.Replace(findText, replaceText + "_Ignore", optionsIgnore);

        // Save the modified document.
        doc.Save("StructuredDocumentTags_Updated.docx");
    }
}
