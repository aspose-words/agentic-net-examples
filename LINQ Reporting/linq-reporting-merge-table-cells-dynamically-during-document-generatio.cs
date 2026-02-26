using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // ---------- Sample data ----------
        var products = new List<Product>
        {
            new Product("Fruits", "Apple", 1.20),
            new Product("Fruits", "Banana", 0.80),
            new Product("Fruits", "Orange", 1.00),
            new Product("Vegetables", "Carrot", 0.50),
            new Product("Vegetables", "Broccoli", 1.30)
        };

        // Order data by Category – this is the LINQ part.
        var ordered = products.OrderBy(p => p.Category).ToList();

        // ---------- Document creation ----------
        Document doc = new Document();                     // create a blank document
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and add a header row.
        Table table = builder.StartTable();

        builder.InsertCell(); builder.Write("Category");
        builder.InsertCell(); builder.Write("Item");
        builder.InsertCell(); builder.Write("Price");
        builder.EndRow();

        // Keep track of the current category to know where a merge should start.
        string currentCategory = null;
        int dataRowIndex = 0; // counts rows *after* the header

        foreach (var p in ordered)
        {
            // First column – write the category only for the first occurrence.
            builder.InsertCell();
            if (p.Category != currentCategory)
            {
                builder.Write(p.Category);
                currentCategory = p.Category;
            }

            // Second column – item name.
            builder.InsertCell();
            builder.Write(p.Name);

            // Third column – price.
            builder.InsertCell();
            builder.Write(p.Price.ToString("C"));

            builder.EndRow();
            dataRowIndex++;
        }

        builder.EndTable();

        // ---------- Apply vertical merge ----------
        // The Category column is column 0. We need to set CellFormat.VerticalMerge for each group.
        for (int i = 0; i < ordered.Count; i++)
        {
            // Is this the first row of a new category group?
            bool isFirstInGroup = i == 0 || ordered[i].Category != ordered[i - 1].Category;
            if (!isFirstInGroup) continue;

            // Determine how many consecutive rows share this category.
            int groupSize = 1;
            for (int j = i + 1; j < ordered.Count; j++)
            {
                if (ordered[j].Category == ordered[i].Category)
                    groupSize++;
                else
                    break;
            }

            // First cell of the group – mark as the start of a merged range.
            builder.MoveToCell(0, i + 1, 0, 0); // table 0, row (header + i), column 0, cell 0
            builder.CellFormat.VerticalMerge = CellMerge.First;

            // Remaining cells – merge to the previous cell.
            for (int k = 1; k < groupSize; k++)
            {
                builder.MoveToCell(0, i + 1 + k, 0, 0);
                builder.CellFormat.VerticalMerge = CellMerge.Previous;
            }
        }

        // ---------- Save the document ----------
        doc.Save("MergedCellsReport.doc"); // DOC format as required
    }

    // Simple POCO representing a product.
    class Product
    {
        public string Category { get; }
        public string Name { get; }
        public double Price { get; }

        public Product(string category, string name, double price)
        {
            Category = category;
            Name = name;
            Price = price;
        }
    }
}
