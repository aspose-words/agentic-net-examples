using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;   // Contains TxtLoadOptions for plain‑text loading

public class Program
{
    public static void Main()
    {
        // Create a sample plain‑text file containing list items with leading whitespaces.
        string txtPath = "sample.txt";
        string[] lines =
        {
            " 1. First item",
            " 2. Second item",
            "   3. Third item with extra spaces",
            "Not a list line",
            " 4. Fourth item"
        };
        File.WriteAllLines(txtPath, lines);

        // Load the text file with DetectNumberingWithWhitespaces enabled so that Aspose.Words
        // automatically recognizes the list items despite the leading spaces.
        TxtLoadOptions loadOptions = new TxtLoadOptions
        {
            DetectNumberingWithWhitespaces = true
        };
        Document doc = new Document(txtPath, loadOptions);

        // Save the resulting document.
        string outPath = "output.docx";
        doc.Save(outPath);
    }
}
