using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source RTF document.
        // The file path can be adjusted as needed.
        Document doc = new Document("SourceDocument.rtf");

        // Use DocumentBuilder to insert section breaks and set the first paragraph
        // of each section to the desired worksheet name.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Ensure we start at the beginning of the document.
        builder.MoveToDocumentStart();

        // --- First worksheet -------------------------------------------------
        // Insert a paragraph that will become the worksheet name.
        builder.Writeln("FirstWorksheet");
        // Add some sample data for this worksheet.
        builder.Writeln("A1\tB1");
        builder.Writeln("A2\tB2");

        // Insert a section break to start a new worksheet.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // --- Second worksheet ------------------------------------------------
        builder.Writeln("SalesReport");
        builder.Writeln("Product\tQuantity");
        builder.Writeln("Widget\t150");
        builder.Writeln("Gadget\t200");

        // Insert another section break if more worksheets are required.
        // builder.InsertBreak(BreakType.SectionBreakNewPage);
        // builder.Writeln("ThirdWorksheet");
        // ...

        // Configure Xlsx save options to create a separate worksheet for each section.
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            SectionMode = XlsxSectionMode.MultipleWorksheets,
            SaveFormat = SaveFormat.Xlsx // Explicitly set the format (optional, default for XlsxSaveOptions)
        };

        // Save the document as an Excel file. Each section becomes a worksheet,
        // and the first paragraph of the section is used as the worksheet name.
        doc.Save("ResultWorkbook.xlsx", xlsxOptions);
    }
}
