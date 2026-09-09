using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph that contains the page number field.
        builder.Write("Page ");
        builder.InsertField("PAGE", ""); // Inserts a PAGE field.
        builder.Writeln(); // Ends the paragraph.

        // Configure the section to display page numbers as uppercase Roman numerals.
        PageSetup pageSetup = doc.FirstSection.PageSetup;
        pageSetup.RestartPageNumbering = true; // Start numbering from the first page.
        pageSetup.PageNumberStyle = NumberStyle.UppercaseRoman;

        // Save the document to a file.
        doc.Save("Output.docx");
    }
}
