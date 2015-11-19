using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Novacode;

namespace TemplNET
{
    public class TemplBuilder
    {
        public bool Debug = TemplConfig.Debug;

        public List<TemplModule> Modules = new List<TemplModule>();
        public List<TemplModule> DefaultModules = new List<TemplModule>();

        public DocX Doc;
        public object Model;
        public string Filename;
        public MemoryStream Stream = new MemoryStream();
        public byte[] Bytes => this.Stream.ToArray();

        private TemplBuilder(bool useDefaultModules)
        {
            //Add default handlers for Sections,tables,pictures,text,TOC,comments. Ordering is important.
            DefaultModules = new List<TemplModule>()
            {
                new TemplSectionModule("Section", this, new string[] { "sec" }),
                new TemplPictureReplaceModule("Picture Replace", this, new string[] { "pic" }),
                new TemplRepeatingTextModule("Repeating Text", this, new string[] { "li" }),
                new TemplRepeatingCellModule("Repeating Cell", this, new string[] { "cel" }),
                new TemplRepeatingRowModule("Repeating Row", this, new string[] { "row" }),
                new TemplPictureReplaceModule("Picture Replace", this, new string[] { "pic" }),
                new TemplTableModule("Table", this, new string[] { "tab" }),
                new TemplPicturePlaceholderModule("Picture Placeholder", this, new string[] { "pic" }),
                new TemplTextModule("Text", this, new string[] { "txt" }),
                new TemplTOCModule("Table of Contents", this, new string[] { "toc" }),
                new TemplCommentsModule("Comments", this, new string[] { "!" }),
            };
            Modules.AddRange(DefaultModules.Where(mod => useDefaultModules));
        }
        /// <summary>
        /// No arguments = generate blank doc
        /// </summary>
        public static DocX EmptyDocument => DocX.Create(new MemoryStream());

        /// <summary>
        /// Start with passed-in doc
        /// </summary>
        /// <param name="document"></param>
        /// <param name="defaultModules"></param>
        public TemplBuilder(DocX document, bool defaultModules = true) : this(defaultModules)
        {
            Doc = document;
            Doc.SaveAs(Stream);
        }

        /// <summary>
        /// Start with file stream
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="defaultModules"></param>
        public TemplBuilder(Stream stream, bool defaultModules = true) : this(defaultModules)
        {
            stream.CopyTo(Stream);
            Doc = DocX.Load(Stream);
        }

        /// <summary>
        /// Start with filename
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="defaultModules"></param>
        public TemplBuilder(string filename, bool defaultModules = true) : this(defaultModules)
        {
            Doc = DocX.Load(filename);
            Doc.SaveAs(Stream);
            Filename = filename;
        }

        /// <summary>
        /// Initialise from byte array
        /// </summary>
        /// <param name="data"></param>
        /// <param name="defaultModules"></param>
        public TemplBuilder(byte[] data, bool defaultModules = true) : this(defaultModules)
        {
            Stream = new MemoryStream(data);
            Doc = DocX.Load(Stream);
        }

        /// <summary>
        /// Save to internal memory. 
        /// Initialising a *new* stream is important due to the internal structure of DocX's load/save.
        /// </summary>
        private void Save()
        {
            Stream = new MemoryStream();
            Doc.SaveAs(Stream);

            if (Debug)
            {
                //this.SaveAs(DebugOutPath);
            }
        }

        /// <summary>
        /// (debug) save to disk
        /// </summary>
        /// <param name="fileName"></param>
        public void SaveAs(string fileName)
        {
            Doc.SaveAs(fileName);
            Filename = fileName;
        }

        /// <summary>
        /// (debug) write a message to the end of the document
        /// </summary>
        /// <param name="s"></param>
        public void Logp(string s)
        {
            if (Debug)
            {
                //Doc.InsertParagraph(s);
            }
        }

        /// <summary>
        /// Main document-building trigger. Generates metadata, runs user-specified DoBuild(), saves document.
        /// </summary>
        /// <returns></returns>
        public void Build()
        {
            Modules.ForEach(mod => mod.Build());
            Save();
        }

        public void Build(object model)
        {
            this.Model = model;
            Build();
        }

    }
}