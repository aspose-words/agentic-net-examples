using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Lists;

class TaskListGenerator
{
    static void Main()
    {
        // Load the XML containing tasks.
        // Example XML structure:
        // <Tasks>
        //   <Task>Buy groceries</Task>
        //   <Task>Call Alice</Task>
        //   <Task>Finish report</Task>
        // </Tasks>
        XDocument xml = XDocument.Load("Tasks.xml");

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a bulleted list using the default bullet template.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletList;
        builder.ListFormat.ListLevelNumber = 0; // Top‑level bullet.

        // Iterate over each <Task> element and add it as a list item.
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
