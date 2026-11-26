using System;
using System.IO;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the first text box shape.
        builder.InsertShape(ShapeType.TextBox, 200, 100);
        Shape firstShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
        SetPropertyIfExists(firstShape.TextBox, "IsLinked", true);
        SetPropertyIfExists(firstShape.TextBox, "IsHead", true); // Mark as head of the linked sequence.

        // Insert the second text box shape.
        builder.InsertShape(ShapeType.TextBox, 200, 100);
        Shape secondShape = (Shape)doc.GetChild(NodeType.Shape, 1, true);
        SetPropertyIfExists(secondShape.TextBox, "IsLinked", true);
        // Leave IsHead at its default value.

        // Save the document.
        string outputPath = "LinkedTextBoxes.docx";
        doc.Save(outputPath);

        // Verify the file was saved.
        Console.WriteLine(File.Exists(outputPath)
            ? $"Document saved successfully to '{outputPath}'."
            : $"Failed to save document to '{outputPath}'.");

        // Check each text box's IsHead property.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            // Ensure the shape actually contains a TextBox.
            if (shape.TextBox != null)
            {
                bool isHead = GetPropertyValue<bool>(shape.TextBox, "IsHead");
                Console.WriteLine($"Shape '{shape.Name}' IsHead: {isHead}");
            }
        }
    }

    // Sets a property by name if it exists on the target object.
    private static void SetPropertyIfExists(object target, string propertyName, object value)
    {
        PropertyInfo prop = target.GetType().GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance);
        if (prop != null && prop.CanWrite)
        {
            prop.SetValue(target, value);
        }
    }

    // Retrieves a property value by name; returns default(T) if the property does not exist.
    private static T GetPropertyValue<T>(object target, string propertyName)
    {
        PropertyInfo prop = target.GetType().GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance);
        if (prop != null && prop.CanRead)
        {
            object val = prop.GetValue(target);
            if (val is T typedVal)
                return typedVal;
        }
        return default;
    }
}
