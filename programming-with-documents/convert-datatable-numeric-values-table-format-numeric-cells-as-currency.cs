using System;
using System.Data;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsDemo
{
    class Program
    {
        static void Main()
        {
            // Sample DataTable with numeric values.
            DataTable table = new DataTable("Sales");
            table.Columns.Add("Product", typeof(string));
            table.Columns.Add("Quantity", typeof(int));
            table.Columns.Add("UnitPrice", typeof(decimal));

            table.Rows.Add("Apple", 10, 0.75m);
            table.Rows.Add("Banana", 5, 0.50m);
            table.Rows.Add("Cherry", 20, 1.20m);

            // Convert the DataTable into a Word table with currency formatting.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a new table.
            Table wordTable = builder.StartTable();

            // ---- Header Row ----
            foreach (DataColumn col in table.Columns)
            {
                builder.InsertCell();
                // Write column name in bold.
                builder.Font.Bold = true;
                builder.Write(col.ColumnName);
                builder.Font.Bold = false;
            }
            builder.EndRow();

            // ---- Data Rows ----
            foreach (DataRow row in table.Rows)
            {
                foreach (DataColumn col in table.Columns)
                {
                    builder.InsertCell();

                    object value = row[col];

                    // If the column type is numeric, format as currency.
                    if (value is IConvertible && IsNumericType(col.DataType))
                    {
                        // Use the current culture's currency format.
                        string formatted = string.Format(CultureInfo.CurrentCulture, "{0:C}", Convert.ToDecimal(value));
                        builder.Write(formatted);
                    }
                    else
                    {
                        builder.Write(value?.ToString() ?? string.Empty);
                    }
                }
                builder.EndRow();
            }

            // End the table.
            builder.EndTable();

            // Save the document.
            doc.Save("DataTableAsCurrencyTable.docx");
        }

        // Helper method to determine if a Type is numeric.
        private static bool IsNumericType(Type type)
        {
            return type == typeof(byte)   || type == typeof(sbyte) ||
                   type == typeof(short)  || type == typeof(ushort) ||
                   type == typeof(int)    || type == typeof(uint) ||
                   type == typeof(long)   || type == typeof(ulong) ||
                   type == typeof(float)  || type == typeof(double) ||
                   type == typeof(decimal);
        }
    }
}
