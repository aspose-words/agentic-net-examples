using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyTableStyleOptions
{
    static void Main()
    {
        // Path to the folder that contains the template (.dot) file.
        string dataDir = @"C:\Data\";

        // Load the DOT template.
        Document doc = new Document(dataDir + "Template.dot");

        // Obtain the first table in the document (assumes at least one table exists).
        Table table = doc.FirstSection.Body.Tables[0];

        // Apply desired style options to the table.
        // Example: enable first row formatting and row banding.
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Save the modified document.
        doc.Save(dataDir + "Result.docx");
    }
}
