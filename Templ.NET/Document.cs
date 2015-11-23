using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Novacode;

namespace TemplNET
{
    public class TemplDoc
    {
        public DocX Doc;
        private string _Filename;
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
        public MemoryStream Stream = new MemoryStream();
        public byte[] Bytes => Stream.ToArray();

        /// <summary>
        /// No arguments = generate blank doc
        /// </summary>
        public static DocX EmptyDocument => DocX.Create(new MemoryStream());

        /// <summary>
        /// Start with passed-in doc
        /// </summary>
        /// <param name="document"></param>
        public TemplDoc(DocX document)
        {
            Doc = document;
            Doc.SaveAs(Stream);
        }

        /// <summary>
        /// Start with file stream
        /// </summary>
        /// <param name="stream"></param>
        public TemplDoc(Stream stream)
        {
            stream.CopyTo(Stream);
            Doc = DocX.Load(Stream);
        }

        /// <summary>
        /// Start with filename
        /// </summary>
        /// <param name="filename"></param>
        public TemplDoc(string filename)
        {
            Filename = filename;
            Doc = DocX.Load(Filename);
            Doc.SaveAs(Stream);
        }

        /// <summary>
        /// Start with file data
        /// </summary>
        /// <param name="data"></param>
        public TemplDoc(byte[] data)
        {
            Stream = new MemoryStream(data);
            Doc = DocX.Load(Stream);
        }

        /// <summary>
        /// Save to internal memory. 
        /// Initialising a *new* stream is important due to the internal structure of DocX's load/save.
        /// </summary>
        public TemplDoc Commit()
        {
            Stream = new MemoryStream();
            Doc.SaveAs(Stream);
            return this;
        }
        /// <summary>
        /// save to disk
        /// </summary>
        /// <param name="fileName"></param>
        public void SaveAs(string fileName)
        {
            // Uses internal getter/setter to filter illegal characters
            Filename = fileName;
            Doc.SaveAs(Filename);
        }
        /// <summary>
        /// clone
        /// </summary>
        public TemplDoc Copy() => new TemplDoc(Commit().Bytes);
    }
}