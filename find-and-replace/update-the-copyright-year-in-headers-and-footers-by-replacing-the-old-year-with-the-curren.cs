using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Define the old year to replace and obtain the current year.
        int oldYear = 2022;
        int newYear = DateTime.Now.Year;

        // Create a sample document with header and footer containing the old year.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln($"© {oldYear} My Company");

        // Footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln($"© {oldYear} My Company");

        // Main body content.
        builder.MoveToDocumentEnd();
        builder.Writeln("This is a sample document.");
        builder.Writeln($"All rights reserved © {oldYear}.");

        // Save the input document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for replacement.
        Document loadedDoc = new Document(inputPath);

        // Prepare a regex that matches the old year as a whole word.
        Regex yearRegex = new Regex(@"\b" + oldYear + @"\b");

        // Perform the replacement across the whole document (including headers/footers).
        int replacedCount = loadedDoc.Range.Replace(yearRegex, newYear.ToString(), new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No year replacements were made.");

        // Save the updated document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);
    }
}
