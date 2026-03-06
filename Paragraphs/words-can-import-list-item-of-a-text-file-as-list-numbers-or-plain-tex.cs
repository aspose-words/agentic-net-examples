using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

class ImportListFromText
{
    static void Main()
    {
        // Path to the source plain‑text file.
        string txtFilePath = @"C:\Data\source.txt";

        // -----------------------------------------------------------------
        // Example 1: Import the text file and let Aspose.Words detect numbered
        // list items (including those delimited by whitespace).
        // -----------------------------------------------------------------
        TxtLoadOptions loadOptionsWithLists = new TxtLoadOptions
        {
            // Detect numbered list items using whitespace as a delimiter.
            DetectNumberingWithWhitespaces = true,
            // Keep automatic numbering detection enabled (default).
            AutoNumberingDetection = true
        };

        // Load the plain‑text file into a Word document using the options above.
        Document docWithLists = new Document(txtFilePath, loadOptionsWithLists);

        // Save the resulting document as DOCX.
        docWithLists.Save(@"C:\Data\Result_WithLists.docx");

        // -----------------------------------------------------------------
        // Example 2: Import the same text file but treat all lines as plain
        // text (no list detection). This is achieved by disabling whitespace
        // based list detection.
        // -----------------------------------------------------------------
        TxtLoadOptions loadOptionsPlain = new TxtLoadOptions
        {
            DetectNumberingWithWhitespaces = false,
            AutoNumberingDetection = true
        };

        Document docPlain = new Document(txtFilePath, loadOptionsPlain);
        docPlain.Save(@"C:\Data\Result_PlainText.docx");
    }
}
