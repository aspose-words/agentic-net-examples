using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class RemoveHtmlTags
{
    static void Main()
    {
        // Load an existing document (replace with your actual file path).
        Document doc = new Document(@"C:\Docs\Input.docx");

        // Regular expression that matches any HTML tag, e.g. <p>, </div>, <img src="..."/>.
        Regex htmlTagPattern = new Regex(@"<[^>]+>", RegexOptions.Compiled);

        // Replace all HTML tags with an empty string.
        // The Replace method works on the whole document range.
        doc.Range.Replace(htmlTagPattern, string.Empty);

        // Save the modified document (replace with your desired output path).
        doc.Save(@"C:\Docs\Output.docx");
    }
}
