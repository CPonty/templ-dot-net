using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Drawing;
using System.IO.Compression;
using Novacode;
using System;

namespace TemplNET
{
    /// <summary>
    /// Generates a debug .zip package of metadata from a document build
    /// </summary>
    /// <example>
    /// <code>
    /// // Normal usage
    /// Templ.Load("C:\template.docx")
    ///      .Build( new { Title = "Hello World!" }, true)
    ///      .SaveAs("C:\output.docx")
    ///      .Debugger
    ///      .SaveAs("C:\debugOutput.zip");
    /// 
    /// // Manual usage
    /// var debugger = new TemplDebugger();
    /// var doc = new TemplDoc("C:\template.docx");
    ///
    /// debugger.AddState(doc, "initial");
    /// // ...
    /// // apply some changes to 'doc'
    /// // ...
    /// debugger.AddState (doc, "middle");
    /// // ...
    /// // apply more changes to 'doc'
    /// // ...
    /// debugger.AddState (doc, "finished");
    /// debugger.AddModuleReport(listOfModulesUsed);
    /// debugger.SaveAs("C:\debugOutput.zip");
    /// </code>
    /// </example>
    public class TemplDebugger
    {
        private string _Filename;
        /// <summary>
        /// <para>Filename (can be null).</para>
        /// <para>On assignment: filters illegal characters and sets the extension to '.zip'</para>
        /// </summary>
        public string Filename
        {
            get
            {
                return _Filename;
            }
            set
            {
                var filename = Path.GetInvalidPathChars().Aggregate(value, (current, c) => current.Replace(c, '-'));
                _Filename = Path.ChangeExtension(filename, ".zip");
            }
        }
        /// <summary>
        /// Debug .zip as memory stream.
        /// <para/>Automatically updates from the underlying documents on Commit(), SaveAs()
        /// </summary>
        public MemoryStream Stream = new MemoryStream();
        /// <summary>
        /// Debug .zip as memory stream.
        /// </summary>
        public byte[] Bytes => Stream.ToArray();

        /// <summary>
        /// Stored document states to include in the .zip package
        /// </summary>
        private IList<TemplDoc> States = new List<TemplDoc>();

        /// <summary>
        /// Stores a named document state
        /// </summary>
        /// <param name="document"></param>
        /// <param name="name"></param>
        public void AddState(TemplDoc document, string name)
        {
            var state = document.Copy();
            state.Filename = $"{States.Count}-{name}";
            States.Add(state);
        }

        /// <summary>
        /// Adds a document containing build statistics to the .zip package
        /// </summary>
        public void AddModuleReport(IEnumerable<TemplModule> modules)
        {
            States.Add(ModuleReport(modules));
        }

        /// <summary>
        /// Generates a document containing build statistics from a set of Modules
        /// </summary>
        /// <param name="modules"></param>
        /// <seealso cref="TemplModule"/>
        /// <seealso cref="Templ.Build"/>
        public TemplDoc ModuleReport(IEnumerable<TemplModule> modules)
        {
            var document = TemplDoc.EmptyDocument;
            var docx = document.Docx;
            var calibri = new FontFamily("Calibri");

            //TODO implement
            Paragraph title = docx.InsertParagraph().Append("Templ.NET Modular Build Report").FontSize(24).Font(calibri).Bold();
            docx.InsertParagraph();
            foreach(TemplModule mod in modules)
            {
                Paragraph heading = docx.InsertParagraph().Append(mod.Name).FontSize(14).Font(calibri).Bold();
                Paragraph body = docx.InsertParagraph().FontSize(10).Font(calibri);
                Table t = body.InsertTableBeforeSelf(1, 2);
                t.Design = TableDesign.ColorfulGridAccent2;
                t.Rows[0].Cells[0].Paragraphs[0].Bold();
                t.Rows[0].Cells[1].Paragraphs[0].Alignment = Alignment.right;
                t.InsertRow(t.Rows[0]);
                t.Rows.Last().Cells[0].Paragraphs[0].InsertText("Match Keys: ");
                t.Rows.Last().Cells[1].Paragraphs[0].InsertText(
                    mod.Prefixes.Select(pre => $"\"{pre}\"").Aggregate((a,b) => $"{a},{b}"));
                t.InsertRow(t.Rows[0]);
                t.Rows.Last().Cells[0].Paragraphs[0].InsertText("Content Matches: ");
                t.Rows.Last().Cells[1].Paragraphs[0].InsertText(String.Format("{0:#,##0}",mod.Statistics.matches));
                t.InsertRow(t.Rows[0]);
                t.Rows.Last().Cells[0].Paragraphs[0].InsertText("Deletions: ");
                t.Rows.Last().Cells[1].Paragraphs[0].InsertText(String.Format("{0:#,##0}",mod.Statistics.removals));
                t.InsertRow(t.Rows[0]);
                t.Rows.Last().Cells[0].Paragraphs[0].InsertText("Elapsed Time: ");
                t.Rows.Last().Cells[1].Paragraphs[0].InsertText(String.Format("{0:#,##0} ms",mod.Statistics.millis));
                foreach (var kv in mod.Statistics.statistics)
                {
                    t.InsertRow(t.Rows[0]);
                    t.Rows.Last().Cells[0].Paragraphs[0].InsertText(kv.Key);
                    t.Rows.Last().Cells[1].Paragraphs[0].InsertText(kv.Value);
                }
                t.Rows[0].Cells[0].Paragraphs[0].InsertText(mod.GetType().ToString().Split('.').Last());
            }

            document.Filename = "ModuleReport";
            document.Commit();
            return document;
        }

        /// <summary>
        /// Builds the .zip file (in memory)
        /// </summary>
        public TemplDebugger Commit()
        {
            Stream = new MemoryStream();
            using (var archive = new ZipArchive(Stream, ZipArchiveMode.Create, true))
            {
                foreach (var state in States)
                {
                    try
                    {
                        var stateFile = archive.CreateEntry(state.Filename);
                        using (var fileStream = stateFile.Open())
                        {
                            state.Stream.Seek(0, SeekOrigin.Begin);
                            state.Stream.CopyTo(fileStream);
                        }
                    }
                    catch
                    {
                        throw new IOException($"Templ: Debugger failed to package states into .zip file. Failed on state \"{state.Filename}\"");
                    }
                }
            }
            return this;
        }

        /// <summary>
        /// Save to disk.
        /// Supplied file name is sanitized and converted to '.zip'
        /// </summary>
        /// <param name="fileName"></param>
        public void SaveAs(string fileName)
        {
            Commit();
            Filename = fileName;
            try
            {
                using (var fileStream = new FileStream(Filename, FileMode.Create))
                {
                    Stream.Seek(0, SeekOrigin.Begin);
                    Stream.CopyTo(fileStream);
                }
            }
            catch
            {
                throw new IOException($"Templ: Debugger failed to write contents to file: \"{Filename}\"");
            }
        }
    }
}
