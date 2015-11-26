using System.Collections.Generic;
using System.Linq;
using System.IO;

namespace TemplNET
{
    /// <summary>
    /// Main class for producing Templ.NET documents
    /// </summary>
    /// <example>
    /// <code>
    /// Templ.Load("C:\template.docx")
    ///      .Build( new { Title = "Hello World!" })
    ///      .SaveAs("C:\output.docx");
    /// 
    /// // With debugging enabled
    /// Templ.Load("C:\template.docx")
    ///      .Build( new { Title = "Hello World!" }, true)
    ///      .SaveAs("C:\output.docx")
    ///      .Debugger
    ///      .SaveAs("C:\debugOutput.zip");
    /// </code>
    /// </example>
    /// <seealso cref="TemplDoc"/>
    public class Templ
    {
        /// <summary>
        /// The .docx MIME type for HTTP responses
        /// </summary>
        public const string DocxMIMEType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
        /// <summary>
        /// Document as bytes
        /// </summary>
        public byte[] Bytes => Document.Bytes;
        /// <summary>
        /// Document as memory stream
        /// </summary>
        public Stream Stream => Document.Stream;
        /// <summary>
        /// Last opened/saved filename of document (can be null)
        /// </summary>
        public string Filename => Document.Filename;

        /// <summary>
        /// Underlying Document instance
        /// </summary>
        private TemplDoc Document;

        /// <summary>
        /// Built-in debugger. Generates a .zip file of debug information at Build time.
        /// </summary>
        /// <example>
        /// <code>
        /// Debugger.SaveAs("C:\debugOutput.zip");
        /// </code>
        /// </example>
        /// <seealso cref="Build(object, bool)"/>
        public TemplDebugger Debugger;

        /// <summary>
        /// Active Modules (text, picture, table etc.). Applied to the document at Build time.
        /// </summary>
        /// <seealso cref="Build(object, bool)"/>
        public List<TemplModule> ActiveModules = new List<TemplModule>();
        public List<string> ModuleNames => ActiveModules.Select(mod => mod.Name).ToList();
        private static List<TemplModule> DefaultModules =>
            new List<TemplModule>()
            {
                new TemplSectionModule("Section"),
                new TemplPictureReplaceModule("Picture Replace"),
                new TemplRepeatingTextModule("Repeating Text"),
                new TemplRepeatingCellModule("Repeating Cell"),
                new TemplRepeatingRowModule("Repeating Row"),
                new TemplPictureReplaceModule("Picture Replace (2nd Pass)"),
                new TemplTableModule("Table"),
                new TemplPicturePlaceholderModule("Picture Placeholder"),
                new TemplHyperlinkModule("Hyperlink"),
                new TemplTextModule("Text"),
                new TemplTOCModule("Table of Contents"),
                new TemplCommentsModule("Comments"),
            };

        private Templ() { } // Empty constructor is private        

        /// <summary>
        /// Initialises a document template, based on the supplied document object
        /// </summary>
        public static Templ Load(TemplDoc document, bool useDefaultModules = true)
        {
            return new Templ()
            {
                Document = document,
                Debugger = new TemplDebugger(),
                ActiveModules = (useDefaultModules ? DefaultModules : new List<TemplModule>())
            };
        }

        /// <summary>
        /// Initialises a document template, based on the supplied document data
        /// </summary>
        public static Templ Load(byte[] templateFile, bool useDefaultModules = true)
        {
            return Load(new TemplDoc(templateFile), useDefaultModules);
        }

        /// <summary>
        /// Initialises a document template, based on the supplied document data stream
        /// </summary>
        public static Templ Load(Stream stream, bool useDefaultModules = true)
        {
            return Load(new TemplDoc(stream), useDefaultModules);
        }

        /// <summary>
        /// Initialises a document template, based on the supplied document file
        /// </summary>
        public static Templ Load(string templatePath, bool useDefaultModules = true)
        {
            return Load(new TemplDoc(templatePath), useDefaultModules);
        }

        /// <summary>
        /// Builds the document, using the provided <paramref name="model"/>.
        /// <para/><see cref="ActiveModules"/> are applied to the document in sequence.
        /// <para/>Enabling <paramref name="debug"/> stores intermediate states after each module.
        /// </summary>
        /// <seealso cref="Debugger"/>
        /// 
        public Templ Build(object model, bool debug = TemplConst.Debug)
        {
            if (debug)
            {
                Debugger.AddState(Document, "Init");
            }
            ActiveModules.ForEach(module =>
            {
                module.Build(Document.Doc, model);
                if (debug && module.Used)
                {
                    Debugger.AddState(Document, module.Name);
                }
            });
            if (debug)
            {
                Debugger.AddModuleReport(ActiveModules);
            }
            Document.Commit();
            Debugger.Commit();
            return this;
        }

        /// <summary>
        /// Saves the document to file.
        /// <para/>Run <see cref="Build"/> first
        /// </summary>
        /// <param name="fileName"></param>
        public Templ SaveAs(string fileName)
        {
            Document.SaveAs(fileName);
            return this;
        }
    }
}