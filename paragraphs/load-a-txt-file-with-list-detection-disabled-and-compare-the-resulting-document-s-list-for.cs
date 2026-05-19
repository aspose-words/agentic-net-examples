using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Prepare sample text containing a list item that uses a whitespace delimiter.
        string sampleText = 
            "Full stop delimiters:\n" +
            "1. First list item 1\n" +
            "2. First list item 2\n" +
            "3. First list item 3\n\n" +
            "Whitespace delimiters:\n" +
            "1 Fourth list item 1\n" +
            "2 Fourth list item 2\n" +
            "3 Fourth list item 3";

        // Create a temporary folder to store the sample TXT file and the output documents.
        string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeExample");
        Directory.CreateDirectory(tempFolder);

        // Write the sample text to a file.
        string txtPath = Path.Combine(tempFolder, "sample.txt");
        File.WriteAllText(txtPath, sampleText);

        // Load the document with list detection enabled (default behavior).
        TxtLoadOptions loadOptionsEnabled = new TxtLoadOptions
        {
            DetectNumberingWithWhitespaces = true
        };
        Document docWithDetection = new Document(txtPath, loadOptionsEnabled);

        // Load the same document with list detection disabled.
        TxtLoadOptions loadOptionsDisabled = new TxtLoadOptions
        {
            DetectNumberingWithWhitespaces = false
        };
        Document docWithoutDetection = new Document(txtPath, loadOptionsDisabled);

        // Helper function to count list items in a document.
        int CountListItems(Document doc) =>
            doc.GetChildNodes(NodeType.Paragraph, true)
               .Cast<Paragraph>()
               .Count(p => p.IsListItem);

        // Count list items and list objects for both documents.
        int listItemsEnabled = CountListItems(docWithDetection);
        int listItemsDisabled = CountListItems(docWithoutDetection);
        int listCountEnabled = docWithDetection.Lists.Count;
        int listCountDisabled = docWithoutDetection.Lists.Count;

        // Output the comparison results.
        Console.WriteLine("List detection enabled:");
        Console.WriteLine($"  List objects: {listCountEnabled}");
        Console.WriteLine($"  List items  : {listItemsEnabled}");
        Console.WriteLine();

        Console.WriteLine("List detection disabled:");
        Console.WriteLine($"  List objects: {listCountDisabled}");
        Console.WriteLine($"  List items  : {listItemsDisabled}");
        Console.WriteLine();

        // Save both documents for manual inspection (optional).
        string outEnabled = Path.Combine(tempFolder, "Detected.docx");
        string outDisabled = Path.Combine(tempFolder, "NotDetected.docx");
        docWithDetection.Save(outEnabled);
        docWithoutDetection.Save(outDisabled);

        // Clean up temporary TXT file (keep the generated DOCX files for review).
        // File.Delete(txtPath);
    }
}
