using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Loading;

class HtmlToWordConverter
{
    static void Main()
    {
        // Path to the source HTML file containing a complex table.
        const string htmlFilePath = @"C:\Temp\ComplexTable.html";

        // Load the HTML document. Use LoadOptions with the LoadFormat property set to Html.
        var loadOptions = new LoadOptions { LoadFormat = LoadFormat.Html };
        Document doc = new Document(htmlFilePath, loadOptions);

        // Iterate over all tables in the document and convert cells that were merged by width
        // into cells merged by the HorizontalMerge flag. This preserves merged cells when saved.
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
            table.ConvertToHorizontallyMergedCells();
        }

        // Save the resulting document as a Word .docx file.
        const string outputFilePath = @"C:\Temp\ConvertedDocument.docx";
        doc.Save(outputFilePath, SaveFormat.Docx);
    }
}
