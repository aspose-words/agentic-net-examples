using System.Text.RegularExpressions;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new document and add some text that uses a custom delimiter (e.g., '|')
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Apple | Banana|Cherry | Date");

        // Regex pattern captures any whitespace before and after the delimiter
        // Group 1 = whitespace before, Group 2 = whitespace after
        Regex pattern = new Regex(@"(\s*)\|(\s*)");

        // Replace the delimiter with a comma, preserving the captured whitespace
        doc.Range.Replace(pattern, "$1,$2");

        // Save the modified document
        doc.Save("Output.docx");
    }
}
