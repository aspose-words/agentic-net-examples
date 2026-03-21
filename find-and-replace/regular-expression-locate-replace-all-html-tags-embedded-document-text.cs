using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class RemoveHtmlTags
{
    static void Main()
    {
        // Create a new document and add some text that contains HTML tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a <b>bold</b> text with <a href='https://example.com'>link</a> and <img src='image.png'/>.");

        // Regular expression that matches any HTML tag, e.g. <p>, </div>, <img src="..."/>.
        Regex htmlTagPattern = new Regex(@"<[^>]+>", RegexOptions.Compiled);

        // Replace all HTML tags with an empty string.
        doc.Range.Replace(htmlTagPattern, string.Empty);

        // Save the modified document to a temporary file and display its path.
        string outputPath = Path.Combine(Path.GetTempPath(), "Output.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
