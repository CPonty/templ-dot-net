using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.IO.Compression;

namespace TemplNET
{
    public class TemplDebugger
    {
        private string _Filename;
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
        //private ZipFile Zip = new ZipFile();
        public MemoryStream Stream = new MemoryStream();
        public byte[] Bytes => Stream.ToArray();
        private IList<TemplDoc> States = new List<TemplDoc>();

        public void AddState(TemplDoc document, string name)
        {
            var state = document.Copy();
            // Use the module name in the state filename
            state.Filename = $"{States.Count}-{name}";
            States.Add(state);
            //Zip.AddEntry(state.Filename, state.Bytes);
        }
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
