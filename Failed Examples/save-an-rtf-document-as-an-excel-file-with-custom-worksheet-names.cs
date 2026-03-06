// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RtfToExcelWithCustomSheets
{
    static void Main()
    {
        // Load the source RTF document.
        Document doc = new Document("input.rtf");

        // Ensure the document has at least one section.
        doc.EnsureMinimum();

        // Set a custom name for the first worksheet (first section).
        doc.Sections[0].Name = "FirstSheet";

        // Add content to the first worksheet.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Data for the first worksheet.");

        // Insert a new section that will become a separate worksheet.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Data for the second worksheet.");

        // Set a custom name for the second worksheet (second section).
        doc.Sections[1].Name = "SecondSheet";

        // Configure save options to create a separate worksheet for each section.
        XlsxSaveOptions saveOptions = new XlsxSaveOptions
        {
            SectionMode = XlsxSectionMode.MultipleWorksheets,
            SaveFormat = SaveFormat.Xlsx
        };

        // Save the document as an Excel file with the custom worksheet names.
        doc.Save("output.xlsx", saveOptions);
    }
}
