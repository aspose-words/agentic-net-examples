using System;
using Aspose.Words;
using Aspose.Words.Tables;

class JoinAdjacentTables
{
    static void Main()
    {
        // Load the document that contains the adjacent tables.
        // Replace the path with the actual location of your source document.
        string inputPath = @"C:\Docs\Tables.docx";
        Document doc = new Document(inputPath);

        // Get the first two tables in the document.
        // It is assumed that the tables are adjacent and we want to merge the second into the first.
        Table firstTable = doc.FirstSection.Body.Tables[0];
        Table secondTable = doc.FirstSection.Body.Tables[1];

        // Move all rows from the second table to the first table.
        // This effectively joins the two tables into one continuous table.
        while (secondTable.HasChildNodes)
        {
            // Append the first row of the second table to the first table.
            firstTable.Rows.Add(secondTable.FirstRow);
        }

        // Remove the now empty second table from the document.
        secondTable.Remove();

        // Save the modified document.
        // Replace the path with the desired output location.
        string outputPath = @"C:\Docs\TablesJoined.docx";
        doc.Save(outputPath);
    }
}
