using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToMhtml
{
    static void Main()
    {
        // Path to the source DOCX file.
        string MyDir = @"C:\Input\";
        // Path to the folder where the MHTML file will be saved.
        string ArtifactsDir = @"C:\Output\";

        // Load an existing document.
        Document doc = new Document(MyDir + "Document.docx");

        // Save the document as MHTML using the SaveFormat enumeration.
        doc.Save(ArtifactsDir + "Document.mht", SaveFormat.Mhtml);

        // Alternatively, use HtmlSaveOptions to control additional MHTML settings.
        // HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Mhtml);
        // doc.Save(ArtifactsDir + "DocumentWithOptions.mht", options);
    }
}
