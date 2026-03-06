// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Cells;          // Used to rename the worksheet and apply Excel‑specific formatting

class DocxToXlsxWithFormatting
{
    static void Main()
    {
        // Paths to the source DOCX and the target XLSX files.
        string dataDir = @"C:\Data\";
        string docPath = Path.Combine(dataDir, "Input.docx");
        string xlsxPath = Path.Combine(dataDir, "Output.xlsx");

        // Load the DOCX document.
        Document doc = new Document(docPath);

        // Configure XlsxSaveOptions.
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            // Export the whole document to a single worksheet.
            SectionMode = XlsxSectionMode.SingleWorksheet,

            // Enable pretty formatting (adds indentation and line breaks where applicable).
            PrettyFormat = true,

            // Explicitly set the save format to XLSX.
            SaveFormat = SaveFormat.Xlsx
        };

        // Save the document as XLSX using the configured options.
        doc.Save(xlsxPath, xlsxOptions);

        // -----------------------------------------------------------------
        // At this point the XLSX file exists, but Aspose.Words does not expose
        // a direct API to set the worksheet name. Use Aspose.Cells to modify
        // the workbook after it has been created.
        // -----------------------------------------------------------------

        // Load the generated XLSX workbook.
        Workbook workbook = new Workbook(xlsxPath);

        // Rename the first (and only) worksheet.
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "MyData";

        // Example of additional Excel formatting:
        // - Set a default column width.
        // - Apply a bold, blue font style to the first row (assumed header row).
        sheet.Cells.StandardWidth = 20; // 20 characters

        Style headerStyle = workbook.CreateStyle();
        headerStyle.Font.IsBold = true;
        headerStyle.Font.Color = System.Drawing.Color.Blue;

        // Apply the style to the first row (row index 0).
        for (int col = 0; col < sheet.Cells.MaxColumn; col++)
        {
            sheet.Cells[0, col].SetStyle(headerStyle);
        }

        // Save the workbook back to the same file (overwrites the previous version).
        workbook.Save(xlsxPath, SaveFormat.Xlsx);
    }
}
