using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string inputPath = "input.docx";
        const string outputPath = "report.pdf";

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX with a date placeholder.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Report generated on {{Date}}.");
        sourceDoc.Save(inputPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Load the DOCX we just created.
        // -----------------------------------------------------------------
        Document doc = new Document(inputPath);

        // -----------------------------------------------------------------
        // 3. Replace all occurrences of the placeholder with the current date.
        // -----------------------------------------------------------------
        string placeholder = "{{Date}}";
        string currentDate = DateTime.Now.ToString("d"); // Short date pattern.

        // Use FindReplaceOptions; only set properties that exist in the current API version.
        FindReplaceOptions replaceOptions = new FindReplaceOptions
        {
            MatchCase = false
            // FindWholeWords property is not available in this version and is therefore omitted.
        };

        doc.Range.Replace(placeholder, currentDate, replaceOptions);

        // -----------------------------------------------------------------
        // 4. Export the updated document to PDF.
        // -----------------------------------------------------------------
        doc.Save(outputPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 5. Verify that the PDF was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The PDF report was not created.");
    }
}
