using System;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Lists;

class TaskListGenerator
{
    static void Main()
    {
        // Sample XML data representing tasks.
        const string xmlContent = @"
<Tasks>
    <Task>Buy groceries</Task>
    <Task>Call Alice</Task>
    <Task>Finish report</Task>
</Tasks>";

        // Load the XML from the string.
        XDocument xml = XDocument.Parse(xmlContent);

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a bulleted list using the default bullet template.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletList;
        builder.ListFormat.ListLevelNumber = 0; // Top‑level bullet.

        // Add each task as a list item.
        foreach (XElement taskElement in xml.Descendants("Task"))
        {
            string taskText = taskElement.Value.Trim();
            if (!string.IsNullOrEmpty(taskText))
            {
                builder.Writeln(taskText);
            }
        }

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the resulting document.
        doc.Save("TaskList.docx");
    }
}
