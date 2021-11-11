using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace TaskTrayMouseShaker
{
    public class MouseController
    {
        [DllImport("User32.dll")]
        private static extern bool SetCursorPos(int X, int Y);

        public static void SetPosition(int x, int y)
        {
            SetCursorPos(x, y);
        }
    }
}
