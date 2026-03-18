using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Markup;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Ole;

class ExportCheckboxesToCsv
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("Input.docx");

        // Prepare a StringBuilder to build CSV content.
        StringBuilder csv = new StringBuilder();
        csv.AppendLine("ControlType,Name,Checked");

        // 1. Process legacy FormField check boxes.
        foreach (FormField field in doc.Range.FormFields)
        {
            if (field.Type == FieldType.FieldFormCheckBox)
            {
                string name = field.Name ?? string.Empty;
                string checkedValue = field.Checked ? "True" : "False";
                csv.AppendLine($"FormField,{Escape(name)},{checkedValue}");
            }
        }

        // 2. Process StructuredDocumentTag (SDT) check boxes.
        NodeCollection sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        foreach (StructuredDocumentTag sdt in sdtNodes)
        {
            if (sdt.SdtType == SdtType.Checkbox)
            {
                // In recent Aspose.Words versions the Checked state is exposed directly.
                // Older versions expose it via sdt.SdtProperties.CheckBox.Checked.
                bool isChecked = sdt.Checked;
                string name = sdt.Title ?? string.Empty;
                string checkedValue = isChecked ? "True" : "False";
                csv.AppendLine($"SDT,{Escape(name)},{checkedValue}");
            }
        }

        // 3. Process Forms 2.0 OLE check box controls.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes)
        {
            if (shape.OleFormat != null && shape.OleFormat.OleControl is CheckBoxControl oleCheckBox)
            {
                string name = shape.Name ?? string.Empty;
                string checkedValue = oleCheckBox.Checked ? "True" : "False";
                csv.AppendLine($"OLE,{Escape(name)},{checkedValue}");
            }
        }

        // Write the CSV data to a file.
        File.WriteAllText("Checkboxes.csv", csv.ToString());
    }

    // Helper to escape commas and quotes in CSV fields.
    private static string Escape(string value)
    {
        if (value.Contains(",") || value.Contains("\""))
        {
            value = value.Replace("\"", "\"\"");
            return $"\"{value}\"";
        }
        return value;
    }
}
