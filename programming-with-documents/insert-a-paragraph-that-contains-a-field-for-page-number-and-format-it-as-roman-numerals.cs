using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph that contains a PAGE field.
        builder.Writeln("Page ");
        builder.InsertField("PAGE", "");

        // Configure the section to display page numbers as uppercase Roman numerals.
        // RestartPageNumbering ensures the numbering starts from the first page.
        PageSetup pageSetup = doc.FirstSection.PageSetup;
        pageSetup.RestartPageNumbering = true;
        pageSetup.PageNumberStyle = NumberStyle.UppercaseRoman;

        // Save the document to the local file system.
        string outputPath = "RomanPageNumbers.docx";
        doc.Save(outputPath);
    }
}
