using System.Collections.Generic;
using System.IO;
using System.Linq;
using Novacode;

namespace TemplNET
{
    /// <summary>
    /// Represents an instance of a document
    /// </summary>
    /// <example>
    /// <code>
    /// TemplDoc.EmptyDocument.SaveAs("C:\empty.docx");
    /// 
    /// var doc = new TemplDoc("C:\empty.docx");
    /// 
    /// var data = doc.Bytes;
    /// 
    /// var doc2 = doc.Copy();
    /// </code>
    /// </example>
    public class TemplDoc
    {
        /// <summary>
        /// Underlying DocX document instance
        /// </summary>
        public DocX Docx;
        /// <summary>
        /// Document as memory stream.
        /// <para/>Automatically updates from the underlying DocX document on Commit(), SaveAs(), Copy()
        /// </summary>
        public MemoryStream Stream = new MemoryStream();
        /// <summary>
        /// Document as bytes
        /// </summary>
        public byte[] Bytes => Stream.ToArray();
        private string _Filename;
        /// <summary>
        /// Filename (can be null).
        /// <para/>On assignment: filters illegal characters and sets the extension to '.docx'
        /// </summary>
        public string Filename
        {
            get
            {
                return _Filename;
            }
            set
            {
                _Filename = Path.GetInvalidPathChars().Aggregate(value, (current, c) => current.Replace(c, '-'));
                _Filename = Path.ChangeExtension(_Filename, ".docx");
            }
        }

        /// <summary>
        /// Generate a blank document
        /// </summary>
        public static TemplDoc EmptyDocument => new TemplDoc(EmptyDocx);
        private static DocX EmptyDocx => DocX.Create(new MemoryStream());

        /// <summary>
        /// Start with passed-in doc
        /// </summary>
        /// <param name="document"></param>
        public TemplDoc(DocX document)
        {
            Docx = document;
            Docx.SaveAs(Stream);
        }

        /// <summary>
        /// New document from data stream
        /// </summary>
        /// <param name="stream"></param>
        public TemplDoc(Stream stream)
        {
            stream.CopyTo(Stream);
            Docx = DocX.Load(Stream);
        }

        /// <summary>
        /// New document from filename
        /// </summary>
        /// <param name="filename"></param>
        public TemplDoc(string filename)
        {
            Filename = filename;
            Docx = DocX.Load(Filename);
            Docx.SaveAs(Stream);
        }

        /// <summary>
        /// New document from file data
        /// </summary>
        /// <param name="data"></param>
        public TemplDoc(byte[] data)
        {
            Stream = new MemoryStream(data);
            Docx = DocX.Load(Stream);
        }

        /// <summary>
        /// Save to internal memory. 
        /// Initialising a *new* stream is important due to the internal structure of DocX's load/save.
        /// </summary>
        public TemplDoc Commit()
        {
            Stream = new MemoryStream();
            Docx.SaveAs(Stream);
            return this;
        }
        /// <summary>
        /// Save to disk.
        /// Supplied file name is sanitized and converted to '.zip'
        /// </summary>
        /// <param name="fileName"></param>
        public void SaveAs(string fileName)
        {
            // Uses internal getter/setter to filter illegal characters
            Filename = fileName;
            Docx.SaveAs(Filename);
        }
        /// <summary>
        /// Clone
        /// </summary>
        public TemplDoc Copy() => new TemplDoc(Commit().Bytes);

        /// <summary>
        /// Retrieves all paragraph references from the document body, header, footer, and table content.
        /// </summary>
        /// <param name="doc"></param>
        /// <returns></returns>
        public static IEnumerable<Paragraph> Paragraphs(DocX doc)
        {
            return doc.Paragraphs
            .Concat(doc.Headers?.odd?.Paragraphs?.ToList() ?? new List<Paragraph>())
            .Concat(doc.Footers?.odd?.Paragraphs?.ToList() ?? new List<Paragraph>());
        }
    }
}