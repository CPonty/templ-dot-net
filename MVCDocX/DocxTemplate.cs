using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using System.Web.Mvc;

namespace DocXMVC
{
    public class DocxTemplate
    {
        private DocxBuilder Builder;

        public byte[] Bytes => Builder.Bytes;
        public Stream Stream => Builder.Stream;
        public string File => Builder.Filename;
        public List<DocxModule> Modules => Builder.Modules;
        public List<string> ModuleNames => Modules.Select(mod => mod.Name).ToList();

        public DocxTemplate(byte[] docxTemplateFile)
        {
            Builder = new DocxBuilder(docxTemplateFile);
        }
        public DocxTemplate(MemoryStream docxTemplateFile)
        {
            Builder = new DocxBuilder(docxTemplateFile);
        }
        public DocxTemplate(string docxTemplatePath)
        {
            Builder = new DocxBuilder(docxTemplatePath);
        }

        public DocxTemplate Build(object model)
        {
            Builder.Build(model);
            return this;
        }
        public FileContentResult Response => Builder.MvcResponse();
    }
}