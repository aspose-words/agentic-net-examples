// Load a Word document.
var doc = new Aspose.Words.Document(@"C:\Input\Sample.docx");

// Configure save options for Excel (XLSX) format.
// The XlsxSaveOptions class is used to control how the document is saved as an Excel workbook.
var xlsxOptions = new Aspose.Words.Saving.XlsxSaveOptions
{
    // Save each Word section to a separate worksheet.
    // This allows you to control the worksheet layout by arranging sections in the source document.
    SectionMode = Aspose.Words.Saving.XlsxSectionMode.MultipleWorksheets,

    // Explicitly set the format to XLSX (optional, the class defaults to this format).
    SaveFormat = Aspose.Words.SaveFormat.Xlsx
};

// Save the document as an Excel workbook.
// Each section of the original Word document will become an individual worksheet.
doc.Save(@"C:\Output\Result.xlsx", xlsxOptions);
