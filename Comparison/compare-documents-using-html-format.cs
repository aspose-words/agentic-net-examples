using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the original and edited HTML documents.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        Document docOriginal = new Document("Original.html", loadOptions);
        Document docEdited = new Document("Edited.html", loadOptions);

        // Ensure both documents have no revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Configure comparison options.
            CompareOptions compareOptions = new CompareOptions
            {
                CompareMoves = false,
                IgnoreFormatting = false,
                Target = ComparisonTargetType.New
            };

            // Perform the comparison; revisions are added to docOriginal.
            docOriginal.Compare(docEdited, "Author", DateTime.Now, compareOptions);
        }

        // Save the comparison result as an HTML file.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            HtmlVersion = HtmlVersion.Html5,
            PrettyFormat = true
        };
        docOriginal.Save("ComparisonResult.html", saveOptions);
    }
}
