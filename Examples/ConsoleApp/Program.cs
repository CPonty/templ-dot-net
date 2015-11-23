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
        static bool debug = true;
        static string root = @"..\..\";

        static Dictionary<string, object> templates
         = new Dictionary<string, object>()
         {
             ["HelloWorld.docx"] = new HelloWorldModel()
        };

        static void Main(string[] args)
        {
            templates.ToList().ForEach(e => FillTemplate(e.Key, e.Value));

            Console.WriteLine("\n All outputs saved to " + Path.GetFullPath(root + ".."));
        }

        static void FillTemplate(string name, object model)
        {
            var file = root + @"Templates\" + name;
            var output = root + @"..\" + name;

            Console.WriteLine("Generating " + name);

            var doc = Templ.Load(file, debug).Build(model).SaveAs(output);
            if (debug) doc.Debugger.SaveAs(output);
        }
    }
}
