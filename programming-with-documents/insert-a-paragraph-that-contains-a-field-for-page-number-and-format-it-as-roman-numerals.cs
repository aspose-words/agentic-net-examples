using System;
using System.IO;
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

        // Insert a paragraph that contains a PAGE field.
        builder.Writeln("Page ");
        // Insert the PAGE field without an initial result; it will be updated later.
        builder.InsertField("PAGE", "");

        // Configure the section to display page numbers as uppercase Roman numerals.
        // RestartPageNumbering ensures numbering starts from 1 for this section.
        PageSetup pageSetup = doc.FirstSection.PageSetup;
        pageSetup.RestartPageNumbering = true;
        pageSetup.PageNumberStyle = NumberStyle.UppercaseRoman;

        // Update all fields so the PAGE field shows the correct value.
        doc.UpdateFields();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "PageNumberRoman.docx");
        doc.Save(outputPath);
    }
}
