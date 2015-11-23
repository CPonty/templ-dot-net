using System.Collections.Generic;
using System.Linq;
using System.IO;

namespace TemplNET
{
    public class Templ
    {
        public const string DocxMIMEType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
        public TemplDoc Document;
        public TemplDebugger Debugger;
        private bool Debug = TemplConst.Debug;

        public byte[] Bytes => Document.Bytes;
        public Stream Stream => Document.Stream;
        public string Filename => Document.Filename;

        public List<TemplModule> Modules = new List<TemplModule>();
        public List<string> ModuleNames => Modules.Select(mod => mod.Name).ToList();
        public static List<TemplModule> DefaultModules =>
             new List<TemplModule>()
            {
                new TemplSectionModule("Section", TemplConst.Tag.Section),
                new TemplPictureReplaceModule("Picture Replace", TemplConst.Tag.Picture),
                new TemplRepeatingTextModule("Repeating Text", TemplConst.Tag.List),
                new TemplRepeatingCellModule("Repeating Cell", TemplConst.Tag.Cell),
                new TemplRepeatingRowModule("Repeating Row", TemplConst.Tag.Row),
                new TemplPictureReplaceModule("Picture Replace", TemplConst.Tag.Picture),
                new TemplTableModule("Table", TemplConst.Tag.Table),
                new TemplPicturePlaceholderModule("Picture Placeholder", TemplConst.Tag.Picture),
                new TemplTextModule("Text", TemplConst.Tag.Text),
                new TemplTOCModule("Table of Contents", TemplConst.Tag.Contents),
                new TemplCommentsModule("Comments", TemplConst.Tag.Comment),
            };

        private Templ() { } // Constructor is private
        public static Templ Load(TemplDoc document, bool debug = TemplConst.Debug, bool useDefaultModules = true)
        {
            return new Templ()
            {
                Document = document,
                Debug = debug,
                Debugger = new TemplDebugger(),
                Modules = (useDefaultModules ? DefaultModules : new List<TemplModule>())
            };
        }
        public static Templ Load(byte[] templateFile, bool debug = TemplConst.Debug, bool useDefaultModules = true)
        {
            return Load(new TemplDoc(templateFile), debug, useDefaultModules);
        }
        public static Templ Load(MemoryStream templateFile, bool debug = TemplConst.Debug, bool useDefaultModules = true)
        {
            return Load(new TemplDoc(templateFile), debug, useDefaultModules);
        }
        public static Templ Load(string templatePath, bool debug = TemplConst.Debug, bool useDefaultModules = true)
        {
            return Load(new TemplDoc(templatePath), debug, useDefaultModules);
        }

        public Templ Build(object model)
        {
            if (Debug)
            {
                Debugger.AddState(Document, "Init");
            }
            Modules.ForEach(module =>
            {
                module.Build(Document.Doc, model);
                if (Debug && module.Used)
                {
                    Debugger.AddState(Document, module.Name);
                }
            });
            Document.Commit();
            Debugger.Commit();
            return this;
        }
        public Templ SaveAs(string fileName)
        {
            Document.SaveAs(fileName);
            return this;
        }
    }
}