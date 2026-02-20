using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the existing DOTX document.
        Document doc = new Document("Template.dotx");

        // Retrieve the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // Apply the desired TableStyleOptions flags.
        // Example: apply formatting to the first row and enable row banding.
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Save the modified document.
        doc.Save("Result.dotx");
    }
}
