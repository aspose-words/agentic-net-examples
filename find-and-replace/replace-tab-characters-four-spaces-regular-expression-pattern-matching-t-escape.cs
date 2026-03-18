using System.Text.RegularExpressions;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("input.docx");

        // Replace every tab character (\t) with four spaces using a regular expression.
        Regex tabPattern = new Regex("\t");
        doc.Range.Replace(tabPattern, "    ");

        // Save the updated document.
        doc.Save("output.docx");
    }
}
