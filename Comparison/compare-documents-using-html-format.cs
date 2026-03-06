using System;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class DocumentComparisonHtml
{
    static void Main()
    {
        // Paths to the original and edited HTML documents.
        string originalPath = "Original.html";
        string editedPath = "Edited.html";

        // Load the HTML documents with default load options.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        Document docOriginal = new Document(originalPath, loadOptions);
        Document docEdited = new Document(editedPath, loadOptions);

        // Ensure both documents have no existing revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Set comparison options as needed (e.g., ignore formatting changes).
            CompareOptions compareOptions = new CompareOptions
            {
                IgnoreFormatting = true,
                Target = ComparisonTargetType.New // Use the edited document as the target.
            };

            // Perform the comparison. Revisions will be added to docOriginal.
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
