using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Novacode;

namespace TemplNET
{
    public class TemplBuilder
    {
        public bool Debug = TemplConst.Debug;

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
                new TemplSectionModule("Section", this, new string[] { TemplConst.Tag.Section }),
                new TemplPictureReplaceModule("Picture Replace", this, new string[] { TemplConst.Tag.Picture }),
                new TemplRepeatingTextModule("Repeating Text", this, new string[] { TemplConst.Tag.List }),
                new TemplRepeatingCellModule("Repeating Cell", this, new string[] { TemplConst.Tag.Cell }),
                new TemplRepeatingRowModule("Repeating Row", this, new string[] { TemplConst.Tag.Row }),
                new TemplPictureReplaceModule("Picture Replace", this, new string[] { TemplConst.Tag.Picture }),
                new TemplTableModule("Table", this, new string[] { TemplConst.Tag.Table }),
                new TemplPicturePlaceholderModule("Picture Placeholder", this, new string[] { TemplConst.Tag.Picture }),
                new TemplTextModule("Text", this, new string[] { TemplConst.Tag.Text }),
                new TemplTOCModule("Table of Contents", this, new string[] { TemplConst.Tag.Contents }),
                new TemplCommentsModule("Comments", this, new string[] { TemplConst.Tag.Comment }),
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
            Filename = filename;
            Doc = DocX.Load(filename);
            Doc.SaveAs(Stream);
        }

        /// <summary>
        /// Start with file data
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
        }

        /// <summary>
        /// save to disk
        /// </summary>
        /// <param name="fileName"></param>
        public void SaveAs(string fileName)
        {
            Filename = fileName;
            Doc.SaveAs(fileName);
        }

        /// <summary>
        /// Create document from template, model
        /// </summary>
        public void Build(object model)
        {
            Model = model;
            Modules.ForEach(mod => mod.Build());
            Save();
        }

    }
}