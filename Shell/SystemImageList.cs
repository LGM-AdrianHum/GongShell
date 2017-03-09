// GongSolutions.Shell - A Windows Shell library for .Net.
// Copyright (C) 2007-2009 Steven J. Kirk
//
// This program is free software; you can redistribute it and/or
// modify it under the terms of the GNU Lesser General Public
// License as published by the Free Software Foundation; either 
// version 2 of the License, or (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU Lesser General Public License for more details.
//
// You should have received a copy of the GNU Lesser General Public 
// License along with this program; if not, write to the Free 
// Software Foundation, Inc., 51 Franklin Street, Fifth Floor,  
// Boston, MA 2110-1301, USA.
//

using System;
using System.Drawing;
using System.Windows.Forms;
using GongSolutions.Shell.Interop;

namespace GongSolutions.Shell
{
    internal class SystemImageList
    {
        private static IntPtr _mSmallImageList;
        private static IntPtr _mLargeImageList;

        private static IntPtr SmallImageList
        {
            get
            {
                if (_mSmallImageList == IntPtr.Zero)
                {
                    InitializeImageLists();
                }
                return _mSmallImageList;
            }
        }

        private static IntPtr LargeImageList
        {
            get
            {
                if (_mLargeImageList == IntPtr.Zero)
                {
                    InitializeImageLists();
                }
                return _mLargeImageList;
            }
        }

        public static void DrawSmallImage(Graphics g, Point point,
            int imageIndex, bool selected)
        {
            var flags = (uint) (imageIndex >> 16);
            var hdc = g.GetHdc();

            try
            {
                if (selected) flags |= (int) ILD.BLEND50;
                ComCtl32.ImageList_Draw(SmallImageList, imageIndex & 0xffff,
                    hdc, point.X, point.Y, flags);
            }
            finally
            {
                g.ReleaseHdc();
            }
        }

        public static void UseSystemImageList(ListView control)
        {
            IntPtr large, small;
            int x, y;

            if (control.LargeImageList == null)
            {
                control.LargeImageList = new ImageList();
            }

            if (control.SmallImageList == null)
            {
                control.SmallImageList = new ImageList();
            }

            Shell32.FileIconInit(true);
            if (!Shell32.Shell_GetImageLists(out large, out small))
            {
                throw new Exception("Failed to get system image list");
            }

            ComCtl32.ImageList_GetIconSize(large, out x, out y);
            control.LargeImageList.ImageSize = new Size(x, y);
            ComCtl32.ImageList_GetIconSize(small, out x, out y);
            control.SmallImageList.ImageSize = new Size(x, y);

            User32.SendMessage(control.Handle, MSG.LVM_SETIMAGELIST,
                (int) LVSIL.LVSIL_NORMAL, LargeImageList);
            User32.SendMessage(control.Handle, MSG.LVM_SETIMAGELIST,
                (int) LVSIL.LVSIL_SMALL, SmallImageList);
        }

        public static void UseSystemImageList(TreeView control)
        {
            User32.SendMessage(control.Handle, MSG.TVM_SETIMAGELIST,
                0, SmallImageList);
        }

        private static void InitializeImageLists()
        {
            Shell32.FileIconInit(true);
            if (!Shell32.Shell_GetImageLists(out _mLargeImageList,
                out _mSmallImageList))
            {
                throw new Exception("Failed to get system image list");
            }
        }
    }
}