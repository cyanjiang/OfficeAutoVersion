using System;
using System.Runtime.InteropServices;

namespace OfficeAutoVersion.Core
{
    /// <summary>
    /// 把文件移入 Windows 回收站（用户可恢复）。调用方在返回 false 时应改为直接删除。
    /// </summary>
    internal static class RecycleBin
    {
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct ShFileOpStruct
        {
            public IntPtr hwnd;
            public uint wFunc;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string pFrom;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string pTo;
            public ushort fFlags;
            [MarshalAs(UnmanagedType.Bool)]
            public bool fAnyOperationsAborted;
            public IntPtr hNameMappings;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string lpszProgressTitle;
        }

        private const uint FoDelete = 0x0003;
        private const ushort FofAllowUndo = 0x0040;
        private const ushort FofNoConfirmation = 0x0010;
        private const ushort FofSilent = 0x0004;

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        private static extern int SHFileOperation(ref ShFileOpStruct fileOp);

        /// <summary>把文件移入回收站，成功返回 true。</summary>
        public static bool Delete(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return false;
            }
            try
            {
                var operation = new ShFileOpStruct
                {
                    wFunc = FoDelete,
                    pFrom = path + "\0\0",
                    fFlags = FofAllowUndo | FofNoConfirmation | FofSilent,
                };
                var result = SHFileOperation(ref operation);
                return result == 0 && !operation.fAnyOperationsAborted;
            }
            catch
            {
                return false;
            }
        }
    }
}
