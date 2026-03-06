using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Tables; // Added for Table class

class FindAndReplaceDemo
{
    static void Main()
    {
        // Load an existing DOCX document.
        // (Assumes the file "Input.docx" exists in the same folder as the executable.)
        Document doc = new Document("Input.docx");

        // -----------------------------------------------------------------
        // 1. Simple find-and-replace across the whole document.
        // Replace the placeholder "_FullName_" with the actual name.
        // The method returns the number of replacements performed.
        // -----------------------------------------------------------------
        int countSimple = doc.Range.Replace("_FullName_", "John Doe");
        Console.WriteLine($"Simple replace count: {countSimple}");

        // -----------------------------------------------------------------
        // 2. Find-and-replace with additional options.
        // Example: replace "Ruby" with "Jade" respecting case sensitivity.
        // -----------------------------------------------------------------
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = true   // Set to false to ignore case.
        };
        int countWithOptions = doc.Range.Replace("Ruby", "Jade", options);
        Console.WriteLine($"Replace with options count: {countWithOptions}");

        // -----------------------------------------------------------------
        // 3. Find-and-replace using a regular expression.
        // Replace every number with a paragraph break.
        // -----------------------------------------------------------------
        int countRegex = doc.Range.Replace(new Regex(@"\d+"), "&p");
        Console.WriteLine($"Regex replace count: {countRegex}");

        // -----------------------------------------------------------------
        // 4. Find-and-replace within a specific node (e.g., a table).
        // This demonstrates how to target a sub‑range instead of the whole document.
        // -----------------------------------------------------------------
        // Ensure the document contains at least one table.
        if (doc.FirstSection.Body.Tables.Count > 0)
        {
            Table table = doc.FirstSection.Body.Tables[0];

            // Replace "Carrots" with "Eggs" only inside the table.
            int tableReplaceCount = table.Range.Replace("Carrots", "Eggs", options);
            Console.WriteLine($"Table replace count: {tableReplaceCount}");
        }

        // -----------------------------------------------------------------
        // 5. Save the modified document.
        // -----------------------------------------------------------------
        doc.Save("Output.docx");
        Console.WriteLine("Document saved as Output.docx");
    }
}
