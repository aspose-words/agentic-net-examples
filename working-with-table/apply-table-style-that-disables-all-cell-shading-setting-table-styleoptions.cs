using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyNoShading
{
    static void Main()
    {
        // Load an existing document that contains a table.
        Document doc = new Document("MyDir/Tables.docx");

        // Get the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // Remove any shading that may be applied directly to the table.
        table.ClearShading();

        // Ensure that no table style shading is applied.
        // Setting StyleOptions to None disables all conditional style formatting,
        // including any shading that could be applied via a style.
        table.StyleOptions = TableStyleOptions.None;

        // Save the modified document.
        doc.Save("ArtifactsDir/Table.NoShading.docx");
    }
}
