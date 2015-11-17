using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using Novacode;

namespace DocXMVC
{
    public class DocxBuilder
    {
        public bool LogEnabled = false;

        public const string MIMEType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
        public static string DebugOutPath = HttpContext.Current.Server.MapPath("/App_Data/out.docx");

        // Sorted dictionary because we want to enforce order + unique names
        public List<DocxModule> Modules = new List<DocxModule>();
        public List<DocxModule> DefaultModules = new List<DocxModule>();

        public object Model;
        public DocX Doc;
        public MemoryStream Stream = new MemoryStream();
        public byte[] Bytes => this.Stream.ToArray();
        public string Filename;

        private DocxBuilder(bool useDefaultModules)
        {
            //Add default handlers for Sections,tables,pictures,text,TOC,comments
            DefaultModules = new List<DocxModule>()
            {
                new DocxSectionModule("Section", this, new string[] { "sec" }),
                new DocxPictureReplaceModule("Picture Replace", this, new string[] { "pic" }),
                new DocxRepeatingTextModule("Repeating Text", this, new string[] { "li" }),
                new DocxRepeatingCellModule("Cell", this, new string[] { "cel" }),
                new DocxRepeatingRowModule("Row", this, new string[] { "row" }),
                new DocxPictureReplaceModule("Picture Replace", this, new string[] { "pic" }),
                new DocxTableModule("Table", this, new string[] { "tab" }),
                new DocxPicturePlaceholderModule("Picture Placeholder", this, new string[] { "pic" }),
                new DocxTextModule("Text", this, new string[] { "txt" }),
                new DocxTOCModule("Table of Contents", this, new string[] { "toc" }),
                new DocxCommentsModule("Comments", this, new string[] { "!" }),
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
        public DocxBuilder(DocX document, bool defaultModules = true) : this(defaultModules)
        {
            Doc = document;
            Doc.SaveAs(Stream);
        }

        /// <summary>
        /// Start with file stream
        /// </summary>
        /// <param name="stream"></param>
        public DocxBuilder(Stream stream, bool defaultModules = true) : this(defaultModules)
        {
            stream.CopyTo(Stream);
            Doc = DocX.Load(Stream);
        }

        /// <summary>
        /// Start with filename
        /// </summary>
        /// <param name="filename"></param>
        public DocxBuilder(string filename, bool defaultModules = true) : this(defaultModules)
        {
            Doc = DocX.Load(filename);
            Doc.SaveAs(Stream);
            Filename = filename;
        }

        /// <summary>
        /// Initialise from byte array
        /// </summary>
        /// <param name="data"></param>
        public DocxBuilder(byte[] data, bool defaultModules = true) : this(defaultModules)
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

            #if DEBUG
            this.SaveAs(DebugOutPath);
            #endif
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
        /// MVC ActionResult to return document as file
        /// </summary>
        /// <returns></returns>
        public FileContentResult MvcResponse()
        {
            return new FileContentResult(this.Bytes, DocxBuilder.MIMEType);
        }

        /// <summary>
        /// (debug) write a message to the end of the document
        /// </summary>
        /// <param name="s"></param>
        public void Logp(string s)
        {
        #if DEBUG
            if (LogEnabled)
            {
                Doc.InsertParagraph(s);
            }
        #endif
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

    public class DocxRegex
    {
        private const RegexOptions Options = RegexOptions.Singleline | RegexOptions.Compiled;

        public string Prefix;
        public Regex Pattern;

        public DocxRegex(string prefix)
        {
            Prefix = prefix;
            Pattern = BuildPattern(prefix);
        }

        public string Text(string name)
        {
            return $"{{{Prefix}:{name}}}";
        }
        public static Regex BuildPattern(string prefix)
        {
            return new Regex(@"\{" + Regex.Escape(prefix) + @":([^\}]+)\}", Options);
        }
    }

    /// <summary>
    /// Used to store/load images & generate (scaled) pictures.
    /// </summary>
    public class DocxGraphic
    {
        protected Novacode.Image Image;
        public MemoryStream Data;
        public double Scalar;
        public Alignment? Alignment;
        private bool Loaded = false;

        public DocxGraphic(Stream data, Alignment? align, double scalar = 1.0)
        {
            Data = new MemoryStream();
            data.CopyTo(Data);
            this.Scalar = scalar;
            this.Alignment = align;
        }
        public DocxGraphic(byte[] data, Alignment? align, double scalar = 1.0)
            : this(new MemoryStream(data), align, scalar)
        { }
        public DocxGraphic(string file, Alignment? align, double scalar = 1.0)
            : this(File.ReadAllBytes(file), align, scalar)
        { }

        public DocxGraphic Load(DocX doc)
        {
            if (!Loaded)
            {
                Image = doc.AddImage(Data);
                Loaded = true;
            }
            return this;
        }

        public Picture Picture(double scalar)
        {
            var pic = Image.CreatePicture();
            pic.Width = (int)(pic.Width * (scalar));
            pic.Height = (int)(pic.Height * (scalar));
            return pic;
        }
        public Picture Picture()
        {
            return Picture(Scalar);
        }
        public Picture Picture(uint width)
        {
            var pic = Image.CreatePicture();
            var scaling = ((float)width / pic.Width);
            pic.Height = (int)(pic.Height * scaling);
            pic.Width = (int)(pic.Width * scaling);
            return pic;
        }
    }
}