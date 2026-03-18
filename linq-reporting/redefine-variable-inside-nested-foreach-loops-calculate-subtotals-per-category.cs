using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample hierarchical data: categories → items → amount.
        var data = new List<Category>
        {
            new Category
            {
                Name = "Fruits",
                Items = new List<Item>
                {
                    new Item { Name = "Apple",  Amount = 3 },
                    new Item { Name = "Banana", Amount = 5 },
                    new Item { Name = "Orange", Amount = 2 }
                }
            },
            new Category
            {
                Name = "Vegetables",
                Items = new List<Item>
                {
                    new Item { Name = "Carrot", Amount = 4 },
                    new Item { Name = "Tomato", Amount = 6 }
                }
            }
        };

        // Build a simple table to display the data.
        builder.StartTable();
        builder.InsertCell(); builder.Write("Category");
        builder.InsertCell(); builder.Write("Item");
        builder.InsertCell(); builder.Write("Amount");
        builder.EndRow();

        // Outer loop – iterate over each category.
        foreach (var category in data)
        {
            // Redefine (reset) the subtotal variable for the current category.
            doc.Variables["Subtotal"] = "0";

            bool firstItem = true;

            // Inner loop – iterate over items belonging to the current category.
            foreach (var item in category.Items)
            {
                builder.InsertCell();

                // Write the category name only for the first row; subsequent rows keep the cell empty.
                if (firstItem)
                {
                    builder.Write(category.Name);
                }

                builder.InsertCell(); builder.Write(item.Name);
                builder.InsertCell(); builder.Write(item.Amount.ToString());

                // Update the subtotal variable with the current item's amount.
                double currentSubtotal = double.Parse(doc.Variables["Subtotal"]);
                currentSubtotal += item.Amount;
                doc.Variables["Subtotal"] = currentSubtotal.ToString();

                builder.EndRow();
                firstItem = false;
            }

            // After processing all items of a category, add a subtotal row.
            builder.InsertCell(); // Empty category cell.
            builder.InsertCell(); builder.Write("Subtotal:");
            builder.InsertCell();

            // Insert a DOCVARIABLE field that will display the final subtotal for the category.
            FieldDocVariable subtotalField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
            subtotalField.VariableName = "Subtotal";

            builder.EndRow();
        }

        builder.EndTable();

        // Ensure all fields (including DOCVARIABLE) are updated to reflect variable values.
        doc.UpdateFields();

        // Save the resulting document.
        doc.Save("NestedForeachSubtotals.docx");
    }

    // Simple POCO classes to hold sample data.
    class Category
    {
        public string Name { get; set; }
        public List<Item> Items { get; set; }
    }

    class Item
    {
        public string Name { get; set; }
        public double Amount { get; set; }
    }
}
