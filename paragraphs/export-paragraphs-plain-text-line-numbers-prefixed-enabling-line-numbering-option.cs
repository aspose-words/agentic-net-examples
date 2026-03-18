using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ExportWithLineNumbers
{
    static void Main()
    {
        // Create a new blank document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add several paragraphs – each will become a separate line in the output.
        for (int i = 1; i <= 10; i++)
            builder.Writeln($"Paragraph {i}");

        // Enable line numbering for the first section.
        // The line numbers will be inserted into the document layout and will appear
        // when the document is saved as plain text.
        PageSetup pageSetup = doc.FirstSection.PageSetup;
        pageSetup.LineStartingNumber = 1;                     // start numbering at 1
        pageSetup.LineNumberCountBy = 1;                      // increment by 1
        pageSetup.LineNumberRestartMode = LineNumberRestartMode.Continuous; // keep numbering across pages
        pageSetup.LineNumberDistanceFromText = 0;             // default distance (no extra spacing)

        // Prepare TxtSaveOptions – the default SaveFormat is Text.
        TxtSaveOptions txtOptions = new TxtSaveOptions();

        // Save the document to a .txt file; line numbers will be prefixed to each line.
        doc.Save("ParagraphsWithLineNumbers.txt", txtOptions);
    }
}
