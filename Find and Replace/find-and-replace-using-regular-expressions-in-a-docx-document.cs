using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace AsposeWordsRegexReplace
{
    class Program
    {
        static void Main()
        {
            // Load an existing DOCX document.
            Document doc = new Document("Input.docx");

            // Define a regular expression to find email addresses.
            Regex emailRegex = new Regex(@"\b[\w\.-]+@[\w\.-]+\.\w{2,4}\b", RegexOptions.IgnoreCase);

            // Set up replace options (optional – here we keep defaults).
            FindReplaceOptions options = new FindReplaceOptions();

            // Replace each email address with the placeholder "[email removed]".
            doc.Range.Replace(emailRegex, "[email removed]", options);

            // Save the modified document.
            doc.Save("Output.docx");
        }
    }
}
