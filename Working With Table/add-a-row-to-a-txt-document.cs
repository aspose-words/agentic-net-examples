using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class AddRowToTxt
{
    static void Main()
    {
        // Path to the source TXT file (if you want to modify an existing document).
        // If the file does not exist, Aspose.Words will create a new empty document.
        string inputPath = "input.txt";

        // Load the TXT document with default load options.
        // TxtLoadOptions can be customized if needed (e.g., encoding, whitespace handling).
        Document doc = new Document(inputPath, new TxtLoadOptions());

        // Create a DocumentBuilder attached to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the end of the document and add a new line (row).
        // Writeln inserts the text followed by a paragraph break, which corresponds to a new row in plain‑text.
        builder.Writeln("This is the added row.");

        // Prepare save options for plain‑text output.
        TxtSaveOptions saveOptions = new TxtSaveOptions();

        // Save the modified document back to TXT format.
        string outputPath = "output.txt";
        doc.Save(outputPath, saveOptions);
    }
}
