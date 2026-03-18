using System.Text.RegularExpressions;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Regular expression that matches the textual copyright symbol "(c)" (case‑insensitive).
        Regex copyrightPattern = new Regex(@"\(c\)", RegexOptions.IgnoreCase);

        // Replace each occurrence with the Unicode copyright character © (U+00A9).
        doc.Range.Replace(copyrightPattern, "\u00A9");

        // Save the updated document.
        doc.Save("Output.docx");
    }
}
