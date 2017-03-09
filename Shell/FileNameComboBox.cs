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
using System.ComponentModel;
using System.IO;
using System.Windows.Forms;

namespace GongSolutions.Shell
{
    /// <summary>
    ///     A filename combo box suitable for use in file Open/Save dialogs.
    /// </summary>
    /// <remarks>
    ///     <para>
    ///         This control extends the <see cref="ComboBox" /> class to provide
    ///         auto-completion of filenames based on the folder selected in a
    ///         <see cref="ShellView" />. The control also automatically navigates
    ///         the ShellView control when the user types a folder path.
    ///     </para>
    /// </remarks>
    public class FileNameComboBox : ComboBox
    {
        private ShellView _mShellView;
        private bool _mTryAutoComplete;

        /// <summary>
        ///     Initializes a new instance of the <see cref="FileNameComboBox" />
        ///     class.
        /// </summary>
        [Category("Behaviour")]
        [DefaultValue(null)]
        public FileFilterComboBox FilterControl { get; set; }

        /// <summary>
        ///     Gets/sets the <see cref="ShellView" /> control that the
        ///     <see cref="FileNameComboBox" /> should look for auto-completion
        ///     hints.
        /// </summary>
        [Category("Behaviour")]
        [DefaultValue(null)]
        public ShellView ShellView
        {
            get { return _mShellView; }
            set
            {
                DisconnectEventHandlers();
                _mShellView = value;
                ConnectEventHandlers();
            }
        }

        /// <summary>
        ///     Occurs when a file name is entered into the
        ///     <see cref="FileNameComboBox" /> and the Return key pressed.
        /// </summary>
        public event EventHandler FileNameEntered;

        /// <summary>
        ///     Determines whether the specified key is a regular input key or a
        ///     special key that requires preprocessing.
        /// </summary>
        /// <param name="keyData">
        ///     One of the <see cref="Keys" /> values.
        /// </param>
        /// <returns>
        ///     true if the specified key is a regular input key; otherwise, false.
        /// </returns>
        protected override bool IsInputKey(Keys keyData)
        {
            if (keyData == Keys.Enter)
            {
                return true;
            }
            return base.IsInputKey(keyData);
        }

        /// <summary>
        ///     Raises the <see cref="Control.KeyDown" /> event.
        /// </summary>
        /// <param name="e">
        ///     A <see cref="KeyEventArgs" /> that contains the event data.
        /// </param>
        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);

            if (e.KeyCode == Keys.Enter)
            {
                if ((Text.Length > 0) && !Open(Text) &&
                    (FilterControl != null))
                {
                    FilterControl.Filter = Text;
                }
            }

            _mTryAutoComplete = false;
        }

        /// <summary>
        ///     Raises the <see cref="Control.KeyPress" /> event.
        /// </summary>
        /// <param name="e">
        ///     A <see cref="KeyPressEventArgs" /> that contains the event data.
        /// </param>
        protected override void OnKeyPress(KeyPressEventArgs e)
        {
            base.OnKeyPress(e);
            _mTryAutoComplete = char.IsLetterOrDigit(e.KeyChar);
        }

        /// <summary>
        ///     Raises the <see cref="Control.TextChanged" /> event.
        /// </summary>
        /// <param name="e">
        ///     An <see cref="EventArgs" /> that contains the event data.
        /// </param>
        protected override void OnTextChanged(EventArgs e)
        {
            base.OnTextChanged(e);
            if (_mTryAutoComplete)
            {
                try
                {
                    AutoComplete();
                }
                catch (Exception)
                {
                    // ignored
                }
            }
        }

        private void AutoComplete()
        {
            var rooted = true;

            if ((Text == string.Empty) ||
                (Text.IndexOfAny(new[] {'?', '*'}) != -1))
            {
                return;
            }

            var path = Path.GetDirectoryName(Text);
            var pattern = Path.GetFileName(Text);

            if (((path == null) || (path == string.Empty)) && (_mShellView != null) &&
                _mShellView.CurrentFolder.IsFileSystem && (_mShellView.CurrentFolder != ShellItem.Desktop))
            {
                path = _mShellView.CurrentFolder.FileSystemPath;
                pattern = Text;
                rooted = false;
            }

            if (path == null) return;
            var matches = Directory.GetFiles(path, pattern + '*');

            for (var n = 0; n < 2; ++n)
            {
                if (matches.Length > 0)
                {
                    var currentLength = Text.Length;
                    Text = rooted ? matches[0] : Path.GetFileName(matches[0]);
                    SelectionStart = currentLength;
                    if (Text != null) SelectionLength = Text.Length;
                    break;
                }
                matches = Directory.GetDirectories(path, pattern + '*');
            }
        }

        private void ConnectEventHandlers()
        {
            if (_mShellView != null)
            {
                _mShellView.SelectionChanged += m_ShellView_SelectionChanged;
            }
        }

        private void DisconnectEventHandlers()
        {
            if (_mShellView != null)
            {
                _mShellView.SelectionChanged -= m_ShellView_SelectionChanged;
            }
        }

        private bool Open(string path)
        {
            var result = false;

            if (File.Exists(path))
            {
                FileNameEntered?.Invoke(this, EventArgs.Empty);
                result = true;
            }
            else if (Directory.Exists(path))
            {
                if (_mShellView != null)
                {
                    _mShellView.Navigate(path);
                    Text = string.Empty;
                    result = true;
                }
            }
            else
            {
                OpenParentOf(path);
                Text = Path.GetFileName(path);
            }

            if (_mShellView != null && !Path.IsPathRooted(path) && _mShellView.CurrentFolder.IsFileSystem)
            {
                result = Open(Path.Combine(_mShellView.CurrentFolder.FileSystemPath,
                    path));
            }

            return result;
        }

        private void OpenParentOf(string path)
        {
            var parent = Path.GetDirectoryName(path);

            if (!string.IsNullOrEmpty(parent) &&
                Directory.Exists(parent))
            {
                _mShellView?.Navigate(parent);
            }
        }

        private void m_ShellView_SelectionChanged(object sender, EventArgs e)
        {
            if ((_mShellView.SelectedItems.Length > 0) &&
                !_mShellView.SelectedItems[0].IsFolder &&
                _mShellView.SelectedItems[0].IsFileSystem)
            {
                Text = Path.GetFileName(_mShellView.SelectedItems[0].FileSystemPath);
            }
        }
    }
}