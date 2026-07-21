using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class ExportParagraphsWithLineNumbers
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a few paragraphs.
        builder.Writeln("First paragraph of the document.");
        builder.Writeln("Second paragraph follows the first one.");
        builder.Writeln("Third paragraph completes the sample.");

        // Enable line numbering for the section via PageSetup.
        PageSetup pageSetup = builder.PageSetup;
        pageSetup.LineStartingNumber = 1;                     // Start numbering at 1.
        pageSetup.LineNumberCountBy = 1;                      // Number every line.
        pageSetup.LineNumberRestartMode = LineNumberRestartMode.Continuous;
        pageSetup.LineNumberDistanceFromText = 0;             // Default distance.

        // Save the document as plain text using TxtSaveOptions.
        TxtSaveOptions saveOptions = new TxtSaveOptions();
        doc.Save("ParagraphsWithLineNumbers.txt", saveOptions);
    }
}
