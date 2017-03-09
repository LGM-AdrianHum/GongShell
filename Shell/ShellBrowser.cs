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
using System.Runtime.InteropServices;
using System.Windows.Forms;
using GongSolutions.Shell.Interop;
using IServiceProvider = GongSolutions.Shell.Interop.IServiceProvider;

namespace GongSolutions.Shell
{
    internal class ShellBrowser : IShellBrowser,
        IOleCommandTarget,
        IServiceProvider
    {
        private StatusBar _mStatusBar;
        protected ShellView MShellView;

        public ShellBrowser(ShellView shellView)
        {
            MShellView = shellView;
        }

        public StatusBar StatusBar
        {
            get { return _mStatusBar; }
            set
            {
                _mStatusBar = value;
                if (_mStatusBar != null)
                {
                    _mStatusBar.ShowPanels = true;
                }
            }
        }

        #region IServiceProvider Members

        HResult IServiceProvider.QueryService(ref Guid guidService,
            ref Guid riid,
            out IntPtr ppvObject)
        {
            if (riid == typeof(IOleCommandTarget).GUID)
            {
                ppvObject = Marshal.GetComInterfaceForObject(this,
                    typeof(IOleCommandTarget));
            }
            else if (riid == typeof(IShellBrowser).GUID)
            {
                ppvObject = Marshal.GetComInterfaceForObject(this,
                    typeof(IShellBrowser));
            }
            else
            {
                ppvObject = IntPtr.Zero;
                return HResult.E_NOINTERFACE;
            }

            return HResult.S_OK;
        }

        #endregion

        #region IShellBrowser Members

        HResult IShellBrowser.GetWindow(out IntPtr phwnd)
        {
            phwnd = MShellView.Handle;
            return HResult.S_OK;
        }

        HResult IShellBrowser.ContextSensitiveHelp(bool fEnterMode)
        {
            return HResult.E_NOTIMPL;
        }

        HResult IShellBrowser.InsertMenusSB(IntPtr intPtrShared, IntPtr lpMenuWidths)
        {
            return HResult.E_NOTIMPL;
        }

        HResult IShellBrowser.SetMenuSB(IntPtr intPtrShared, IntPtr holemenuRes, IntPtr intPtrActiveObject)
        {
            return HResult.E_NOTIMPL;
        }

        HResult IShellBrowser.RemoveMenusSB(IntPtr intPtrShared)
        {
            return HResult.E_NOTIMPL;
        }

        HResult IShellBrowser.SetStatusTextSB(IntPtr pszStatusText)
        {
            if (_mStatusBar != null)
            {
                _mStatusBar.Panels.Clear();
                _mStatusBar.Panels.Add(Marshal.PtrToStringUni(pszStatusText));
            }
            return HResult.S_OK;
        }

        HResult IShellBrowser.EnableModelessSB(bool fEnable)
        {
            return HResult.E_NOTIMPL;
        }

        HResult IShellBrowser.TranslateAcceleratorSB(IntPtr pmsg, ushort wId)
        {
            return HResult.E_NOTIMPL;
        }

        HResult IShellBrowser.BrowseObject(IntPtr pidl, SBSP wFlags)
        {
            if ((wFlags & SBSP.SBSP_RELATIVE) != 0)
            {
                // ReSharper disable once UnusedVariable
                var shellItem = new ShellItem(MShellView.CurrentFolder, pidl);
            }
            else if ((wFlags & SBSP.SBSP_PARENT) != 0)
            {
                MShellView.NavigateParent();
            }
            else if ((wFlags & SBSP.SBSP_NAVIGATEBACK) != 0)
            {
                MShellView.NavigateBack();
            }
            else if ((wFlags & SBSP.SBSP_NAVIGATEFORWARD) != 0)
            {
                MShellView.NavigateForward();
            }
            else
            {
                MShellView.Navigate(new ShellItem(ShellItem.Desktop, pidl));
            }
            return HResult.S_OK;
        }

        HResult IShellBrowser.GetViewStateStream(uint grfMode, IntPtr ppStrm)
        {
            return HResult.E_NOTIMPL;
        }

        HResult IShellBrowser.GetControlWindow(FCW id, out IntPtr lpIntPtr)
        {
            if ((id == FCW.FCW_STATUS) && (_mStatusBar != null))
            {
                lpIntPtr = _mStatusBar.Handle;
                return HResult.S_OK;
            }
            lpIntPtr = IntPtr.Zero;
            return HResult.E_NOTIMPL;
        }

        HResult IShellBrowser.SendControlMsg(FCW id, MSG uMsg, uint wParam,
            uint lParam, IntPtr pret)
        {
            var result = 0;

            if ((id == FCW.FCW_STATUS) && (_mStatusBar != null))
            {
                result = User32.SendMessage(_mStatusBar.Handle,
                    uMsg, (int) wParam, (int) lParam);
            }

            if (pret != IntPtr.Zero)
            {
                Marshal.WriteInt32(pret, result);
            }

            return HResult.S_OK;
        }

        HResult IShellBrowser.QueryActiveShellView(out IShellView ppshv)
        {
            ppshv = null;
            return HResult.E_NOTIMPL;
        }

        HResult IShellBrowser.OnViewWindowActive(IShellView ppshv)
        {
            return HResult.E_NOTIMPL;
        }

        HResult IShellBrowser.SetToolbarItems(IntPtr lpButtons, uint nButtons, uint uFlags)
        {
            return HResult.E_NOTIMPL;
        }

        #endregion

        #region IOleCommandTarget Members

        void IOleCommandTarget.QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, ref OLECMDTEXT cmdText)
        {
            throw new NotImplementedException("The method or operation is not implemented.");
        }

        void IOleCommandTarget.Exec(ref Guid pguidCmdGroup, uint nCmdId, uint nCmdExecOpt, ref object pvaIn,
            ref object pvaOut)
        {
            throw new NotImplementedException("The method or operation is not implemented.");
        }

        #endregion
    }

    internal class DialogShellBrowser : ShellBrowser, ICommDlgBrowser
    {
        public DialogShellBrowser(ShellView shellView)
            : base(shellView)
        {
        }

        #region ICommDlgBrowser Members

        HResult ICommDlgBrowser.OnDefaultCommand(IShellView ppshv)
        {
            var selected = MShellView.SelectedItems;

            if ((selected.Length > 0) && selected[0].IsFolder)
            {
                try
                {
                    MShellView.Navigate(selected[0]);
                }
                catch (Exception)
                {
                    // ignored
                }
            }
            else
            {
                MShellView.OnDoubleClick(EventArgs.Empty);
            }

            return HResult.S_OK;
        }

        HResult ICommDlgBrowser.OnStateChange(IShellView ppshv, CDBOSC uChange)
        {
            if (uChange == CDBOSC.CDBOSC_SELCHANGE)
            {
                MShellView.OnSelectionChanged();
            }
            return HResult.S_OK;
        }

        HResult ICommDlgBrowser.IncludeObject(IShellView ppshv, IntPtr pidl)
        {
            return MShellView.IncludeItem(pidl)
                ? HResult.S_OK
                : HResult.S_FALSE;
        }

        #endregion
    }
}