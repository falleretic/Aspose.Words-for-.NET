using System.Collections.Generic;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    //ExStart:Manager
    public class Manager
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public byte[] Photo { get; set; }
        public IEnumerable<Contract> Contracts { get; set; }
    }
    //ExEnd:Manager
}