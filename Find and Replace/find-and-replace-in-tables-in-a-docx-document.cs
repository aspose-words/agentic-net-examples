using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Replacing;
using System.Text.RegularExpressions;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Text to find and its replacement.
        string findText = "OldValue";
        string replaceText = "NewValue";

        // Optional find/replace options.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = false
        };

        // Iterate over all tables in the document.
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
            // Iterate over all cells of the current table.
            foreach (Cell cell in table.GetChildNodes(NodeType.Cell, true))
            {
                // Perform a simple string replace inside the cell's range.
                cell.Range.Replace(findText, replaceText, options);

                // If a regular expression is required, use the overload below:
                // cell.Range.Replace(new Regex(findText), replaceText, options);
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
