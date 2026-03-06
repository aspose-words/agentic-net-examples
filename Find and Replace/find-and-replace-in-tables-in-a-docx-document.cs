using System;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Tables; // <-- added namespace for Table

class TableFindReplace
{
    static void Main()
    {
        // Load the DOCX document (lifecycle rule: load)
        Document doc = new Document("Input.docx");

        // Define find/replace options (case‑sensitive, whole‑word only)
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = true,
            FindWholeWordsOnly = true
        };

        // Text to find and its replacement
        string findText = "Carrots";
        string replaceText = "Eggs";

        // Iterate through all tables in the document
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
            // Perform replace on the whole table range
            table.Range.Replace(findText, replaceText, options);
        }

        // Save the modified document (lifecycle rule: save)
        doc.Save("Output.docx");
    }
}
