using System;
using System.Collections.Generic;
using System.Text;

namespace Import.CrystalReports
{
    public interface IStiTreeLog
    {
        void OpenLog(string headerMessage);
        void CloseLog();
        void OpenNode(string message);
        void WriteNode(string message);
        void WriteNode(string message, object arg);
        void CloseNode();
    }
}
