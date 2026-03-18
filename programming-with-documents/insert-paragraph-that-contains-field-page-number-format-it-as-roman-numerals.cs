using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RomanPageNumberExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a new paragraph.
        builder.Writeln();

        // Write some text before the page number field (optional).
        builder.Write("Page ");

        // Insert a PAGE field that will display the current page number.
        // The overload with two parameters inserts the field without an initial result.
        builder.InsertField("PAGE", "");

        // Configure the document to display page numbers as uppercase Roman numerals.
        // Apply the setting to the first (and only) section.
        Section firstSection = doc.FirstSection;
        firstSection.PageSetup.PageNumberStyle = NumberStyle.UppercaseRoman;
        firstSection.PageSetup.RestartPageNumbering = true; // start numbering from 1

        // Save the document.
        doc.Save("RomanPageNumbers.docx");
    }
}
