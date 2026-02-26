using System;
using System.Drawing;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class FindReplaceWithHighlight
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add some sample text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document demonstrates Aspose find and replace functionality.");
        builder.Writeln("Aspose allows developers to work with Word documents programmatically.");

        // Configure FindReplaceOptions to apply highlighting to the replacement text.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ApplyFont.HighlightColor = Color.Yellow; // Highlight color for the new text.

        // Perform the find-and-replace operation.
        // Replace the word "Aspose" with "Aspose.Words" and apply the highlighting.
        doc.Range.Replace("Aspose", "Aspose.Words", options);

        // Save the modified document to disk.
        doc.Save("FindReplaceWithHighlight.docx");
    }
}
