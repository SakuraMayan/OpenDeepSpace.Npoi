using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenDeepSpace.Npoi
{
    /// <summary>
    /// Npoi内存流 继承MemoryStream 当workbook关闭时自己来决定是否关闭它
    /// </summary>
    public class NpoiMemoryStream:MemoryStream
    {
        public NpoiMemoryStream()
        {

        }

        public bool AllowClose { get; set; }

        public override void Close()
        {
            if (AllowClose)
                base.Close();
        }
    }
}
