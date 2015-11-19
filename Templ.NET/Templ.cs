using System.Collections.Generic;
using System.Linq;
using System.IO;

namespace TemplNET
{
    public class Templ
    {
        public const string DocxMIMEType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
        private TemplBuilder Builder;

        public byte[] Bytes => Builder.Bytes;
        public Stream Stream => Builder.Stream;
        public string Filename => Builder.Filename;
        public List<TemplModule> Modules => Builder.Modules;
        public List<string> ModuleNames => Modules.Select(mod => mod.Name).ToList();

        public Templ(byte[] templateFile)
        {
            Builder = new TemplBuilder(templateFile);
        }
        public Templ(MemoryStream templateFile)
        {
            Builder = new TemplBuilder(templateFile);
        }
        public Templ(string templatePath)
        {
            Builder = new TemplBuilder(templatePath);
        }

        public Templ WithDebugging(bool debug)
        {
            Builder.Debug = debug;
            return this;
        }
        public Templ Build(object model)
        {
            Builder.Build(model);
            return this;
        }
        public Templ SaveAs(string fileName)
        {
            Builder.SaveAs(fileName);
            return this;
        }

    }
}