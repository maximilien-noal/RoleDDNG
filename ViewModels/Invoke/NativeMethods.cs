using RoleDDNG.Models.Structs;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace RoleDDNG.ViewModels.Invoke
{
    internal class NativeMethods
    {
        private static Encoding encoding = new UTF8Encoding();

        [DllImport("user32.dll")]
        private static extern bool SetWindowPlacement(IntPtr hWnd, [In] ref WindowPlacement lpwndpl);

        [DllImport("user32.dll")]
        private static extern bool GetWindowPlacement(IntPtr hWnd, out WindowPlacement lpwndpl);

        private const int SW_SHOWNORMAL = 1;
        private const int SW_SHOWMINIMIZED = 2;

        public static void SetPlacement(IntPtr windowHandle, string placementXml)
        {
            if (string.IsNullOrEmpty(placementXml))
            {
                return;
            }

            WindowPlacement placement;
            byte[] xmlBytes = encoding.GetBytes(placementXml);

            using (MemoryStream memoryStream = new MemoryStream(xmlBytes))
            {
                XmlReaderSettings settings = new XmlReaderSettings() { XmlResolver = null };
                using (XmlReader reader = XmlReader.Create(memoryStream, settings))
                {
                    var serializer = new XmlSerializer(typeof(WindowPlacement));
                    placement = (WindowPlacement)serializer.Deserialize(reader);
                }
            }

            placement.Length = Marshal.SizeOf(typeof(WindowPlacement));
            placement.Flags = 0;
            placement.ShowCmd = (placement.ShowCmd == SW_SHOWMINIMIZED ? SW_SHOWNORMAL : placement.ShowCmd);
            SetWindowPlacement(windowHandle, ref placement);
        }

        public static string GetPlacement(IntPtr windowHandle)
        {
            WindowPlacement placement = new WindowPlacement();
            GetWindowPlacement(windowHandle, out placement);

            using (MemoryStream memoryStream = new MemoryStream())
            {
                using (XmlTextWriter xmlTextWriter = new XmlTextWriter(memoryStream, Encoding.UTF8))
                {
                    new XmlSerializer(typeof(WindowPlacement)).Serialize(xmlTextWriter, placement);
                    byte[] xmlBytes = memoryStream.ToArray();
                    return encoding.GetString(xmlBytes);
                }
            }
        }
    }
}