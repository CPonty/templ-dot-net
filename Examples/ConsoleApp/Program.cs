using System;
using System.Linq;
using System.Collections.Generic;
using TemplNET;
using System.IO;

/*
    This console app uses a local reference to the debug build of "Templ.dll".
    Build the accompanying "Templ.NET" project first!

    To use the release build:
     - Remove References "DocX", "Templ"
     - Run in Package Manager console:
        Install-Package Templ.NET
*/
namespace TemplTest
{
    /// <summary>
    /// <para>Simple console application to run all the example templates</para>
    /// </summary>
    class Program
    {
        static string root = @"..\..\..\"; // "Examples" folder

        static Dictionary<string, object> templates
         = new Dictionary<string, object>()
         {
             ["HelloWorld.docx"] = new HelloWorldModel()
        };

        static void Main(string[] args)
        {
            templates.ToList().ForEach(e => FillTemplate(e.Key, e.Value));

            Console.WriteLine("\n All outputs saved to " + Path.GetFullPath(root));
        }

        static void FillTemplate(string file, object model)
        {
            var filePath = root + @"Templates\" + file;
            var output = root + file;

            Console.WriteLine("Generating " + file);

            Templ.Load(filePath).Build(model).SaveAs(output);
        }
    }
}
