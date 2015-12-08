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
    public class Program
    {
        public static string root = @"..\..\";
        public static bool debug = true;
        public static HandleFailAction modelEntryFailAction = HandleFailAction.exception;

        static void Main(string[] args)
        {
            Generate("HelloWorld.docx", new HelloWorldModel());
            Generate("ContentDemo.docx", new ContentDemoModel());

            Console.WriteLine("\n All outputs saved to " + Path.GetFullPath(root + ".."));

            System.Diagnostics.Process.Start(root + @"..\" + "ContentDemo.docx");
        }

        static void Generate(string name, object model)
        {
            var file = root + @"Templates\" + name;
            var output = root + @"..\" + name;

            Console.WriteLine("Generating " + name);

            var doc = Templ.Load(file).Build(model, debug, modelEntryFailAction).SaveAs(output);
            if (debug) doc.Debugger.SaveAs(output);
        }
    }
}
