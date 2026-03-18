using System;
using System.Text.RegularExpressions;
using Aspose.Words;

class ReplaceEmDash
{
    static void Main()
    {
        // Load an existing document (replace "input.docx" with your file path)
        Document doc = new Document("input.docx");

        // Regular expression that matches the Unicode em dash character (U+2014)
        Regex emDashPattern = new Regex("\u2014");

        // Replace all em dashes with a standard hyphen (U+002D)
        doc.Range.Replace(emDashPattern, "-");

        // Save the modified document (replace "output.docx" with your desired output path)
        doc.Save("output.docx");
    }
}
