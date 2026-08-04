using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a paragraph that contains the text "Page " followed by a PAGE field.
        builder.Write("Page ");
        // Insert a PAGE field; the second argument is the field result placeholder (empty string).
        builder.InsertField("PAGE", "");
        // End the paragraph.
        builder.Writeln();

        // Configure the section to display page numbers as uppercase Roman numerals.
        // Apply the setting to the first (and only) section of the document.
        doc.FirstSection.PageSetup.PageNumberStyle = NumberStyle.UppercaseRoman;
        // Ensure numbering starts at 1 and restarts for this section.
        doc.FirstSection.PageSetup.RestartPageNumbering = true;
        doc.FirstSection.PageSetup.PageStartingNumber = 1;

        // Save the document to a file in the current working directory.
        doc.Save("RomanPageNumbers.docx");
    }
}
