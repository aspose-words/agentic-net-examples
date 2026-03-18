using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class UpdateProductNames
{
    static void Main()
    {
        // Load the existing document.
        Document doc = new Document("Input.docx");

        // Define the old and new product names.
        string oldProductName = "AcmeWidget";
        string newProductName = "AcmeGadget";

        // Configure find/replace options:
        // - Ignore case (MatchCase = false)
        // - Replace only whole words (FindWholeWordsOnly = true)
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = true
        };

        // Perform the replacement across the entire document.
        doc.Range.Replace(oldProductName, newProductName, options);

        // Save the updated document.
        doc.Save("Output.docx");
    }
}
