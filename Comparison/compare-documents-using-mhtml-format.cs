using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Paths to the original and edited MHTML documents.
        string originalPath = @"C:\Docs\Original.mht";
        string editedPath   = @"C:\Docs\Edited.mht";

        // Load the original document from MHTML.
        LoadOptions loadOriginal = new LoadOptions();
        loadOriginal.LoadFormat = LoadFormat.Mhtml;
        Document docOriginal = new Document(originalPath, loadOriginal);

        // Load the edited document from MHTML.
        LoadOptions loadEdited = new LoadOptions();
        loadEdited.LoadFormat = LoadFormat.Mhtml;
        Document docEdited = new Document(editedPath, loadEdited);

        // Ensure both documents have no revisions before comparison.
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // Set up comparison options (optional – customize as needed).
        CompareOptions compareOptions = new CompareOptions
        {
            // Example: track changes at the word level and ignore formatting.
            Granularity = Granularity.WordLevel,
            IgnoreFormatting = true,
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. The original document will receive revision marks.
        docOriginal.Compare(docEdited, "JD", DateTime.Now, compareOptions);

        // Save the comparison result as an MHTML document.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // Use CID URLs for resources to improve compatibility with some mail agents.
            ExportCidUrlsForMhtmlResources = true,
            // Choose HTML version; Xhtml is commonly used for MHTML.
            HtmlVersion = HtmlVersion.Xhtml,
            // Keep the output readable.
            PrettyFormat = true
        };

        string resultPath = @"C:\Docs\ComparisonResult.mht";
        docOriginal.Save(resultPath, saveOptions);
    }
}
