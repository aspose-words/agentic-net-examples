using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToExcel
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document(MyDir + "Document.docx");

        // Create XlsxSaveOptions appropriate for the Xlsx format.
        XlsxSaveOptions xlsxOptions = (XlsxSaveOptions)SaveOptions.CreateSaveOptions(SaveFormat.Xlsx);
        // Optional: save each section of the Word document to a separate worksheet.
        xlsxOptions.SectionMode = XlsxSectionMode.MultipleWorksheets;

        // Save the document as an Excel file.
        doc.Save(ArtifactsDir + "Document.ConvertToXlsx.xlsx", xlsxOptions);
    }

    // Replace these with your actual input and output directories.
    private static readonly string MyDir = @"C:\Input\";
    private static readonly string ArtifactsDir = @"C:\Output\";
}
