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
using System.Drawing;
using System.Windows.Forms;

namespace GongSolutions.Shell
{
    /// <summary>
    ///     Provides a toolbar showing common places in the computer's
    ///     filesystem.
    /// </summary>
    /// <remarks>
    ///     Use the <see cref="PlacesToolbar" /> control to display a
    ///     toolbar listing common places in the computer's filesystem,
    ///     similar to that on the far left-side of the standard file
    ///     open/save dialogs.
    /// </remarks>
    public partial class PlacesToolbar : UserControl
    {
        /// <summary>
        ///     Initializes a new instance of the
        ///     <see cref="PlacesToolbar" /> class.
        /// </summary>
        public PlacesToolbar()
        {
            InitializeComponent();
            // ReSharper disable once VirtualMemberCallInConstructor
            AutoSize = true;
            toolStrip.Renderer = new Renderer();
        }

        /// <summary>
        ///     A <see cref="C:ShellView" /> that should be automatically
        ///     navigated when the user clicks on an entry in the toolbar.
        /// </summary>
        [DefaultValue(null)]
        public ShellView ShellView { get; set; }

        /// <summary>
        ///     Adds a new folder to the toolbar.
        /// </summary>
        /// <param name="folder">
        ///     A <see cref="ShellItem" /> representing the folder to be added.
        /// </param>
        /// <returns>
        ///     true if the item was sucessfully added, false otherwise.
        /// </returns>
        public bool Add(ShellItem folder)
        {
            var include = IncludeItem(folder);

            if (include)
            {
                var button = new ToolStripButton();
                Image image = folder.ShellIcon.ToBitmap();

                toolStrip.ImageScalingSize = image.Size;
                button.AutoSize = false;
                button.Image = image;
                button.Size = new Size(84, image.Height + 35);
                button.Tag = folder;
                button.Text = WrapButtonText(button, folder.DisplayName);
                button.TextImageRelation = TextImageRelation.ImageAboveText;
                button.ToolTipText = folder.ToolTipText;
                button.Click += button_Click;
                toolStrip.Items.Add(button);
            }

            return include;
        }

        /// <summary>
        ///     Occurs when the <see cref="PlacesToolbar" /> control wants to
        ///     know if it should include an item.
        /// </summary>
        /// <remarks>
        ///     This event allows the items displayed in the
        ///     <see cref="PlacesToolbar" /> control to be filtered.
        /// </remarks>
        public event FilterItemEventHandler FilterItem;

        /// <summary>
        ///     Overrides the <see cref="Control.OnVisibleChanged" /> method.
        /// </summary>
        /// <param name="e" />
        protected override void OnVisibleChanged(EventArgs e)
        {
            base.OnVisibleChanged(e);
            if (toolStrip.Items.Count == 0) CreateDefaultItems();
        }

        private void CreateDefaultItems()
        {
            const int csidlNetwork = 0x0012;
            Add(new ShellItem(Environment.SpecialFolder.Recent));
            Add(new ShellItem(Environment.SpecialFolder.Desktop));
            Add(new ShellItem(Environment.SpecialFolder.Personal));
            Add(new ShellItem(Environment.SpecialFolder.MyComputer));
            Add(new ShellItem((Environment.SpecialFolder) csidlNetwork));
        }

        private bool IncludeItem(ShellItem item)
        {
            if (FilterItem != null)
            {
                var e = new FilterItemEventArgs(item);
                FilterItem(this, e);
                return e.Include;
            }
            return true;
        }

        private string WrapButtonText(ToolStripButton button, string s)
        {
            var g = Graphics.FromHwnd(toolStrip.Handle);

            if (g.MeasureString(s, button.Font).Width >
                button.ContentRectangle.Width)
            {
                var lastSpace = s.LastIndexOf(' ');
                return s.Substring(0, lastSpace) + "\r\n" +
                       s.Substring(lastSpace + 1);
            }

            return s;
        }

        private void button_Click(object sender, EventArgs e)
        {
            if (ShellView != null)
            {
                var folder = (ShellItem) ((ToolStripButton) sender).Tag;
                ShellView.Navigate(folder);
            }
        }

        private class Renderer : ToolStripSystemRenderer
        {
            protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e)
            {
                var rect = new Rectangle(0, 0,
                    e.ToolStrip.Width - 1, e.ToolStrip.Height - 1);

                if (Application.RenderWithVisualStyles)
                {
                    ControlPaint.DrawVisualStyleBorder(e.Graphics, rect);
                }
                else
                {
                    ControlPaint.DrawBorder3D(e.Graphics, rect,
                        Border3DStyle.Sunken);
                }
            }
        }

        #region Hidden Properties

        /// <summary>
        ///     This property does not apply to the
        ///     <see cref="PlacesToolbar" /> class.
        /// </summary>
        [Browsable(false)]
        public override bool AllowDrop
        {
            get { return base.AllowDrop; }
            set { base.AllowDrop = value; }
        }

        /// <summary>
        ///     This property is not relevant to the
        ///     <see cref="PlacesToolbar" /> class.
        /// </summary>
        [Browsable(false)]
        public override bool AutoScroll
        {
            get { return base.AutoScroll; }
            set { base.AutoScroll = value; }
        }

        /// <summary>
        ///     This property is not relevant to the
        ///     <see cref="PlacesToolbar" /> class.
        /// </summary>
        [Browsable(false)]
        public new Size AutoScrollMargin
        {
            get { return base.AutoScrollMargin; }
            set { base.AutoScrollMargin = value; }
        }

        /// <summary>
        ///     This property is not relevant to the
        ///     <see cref="PlacesToolbar" /> class.
        /// </summary>
        [Browsable(false)]
        public new Size AutoScrollMinSize
        {
            get { return base.AutoScrollMinSize; }
            set { base.AutoScrollMinSize = value; }
        }

        /// <summary>
        ///     This property is not relevant to the
        ///     <see cref="PlacesToolbar" /> class.
        /// </summary>
        [Browsable(false), DefaultValue(true)]
        public override bool AutoSize
        {
            get { return base.AutoSize; }
            set { base.AutoSize = value; }
        }

        /// <summary>
        ///     This property is not relevant to the
        ///     <see cref="PlacesToolbar" /> class.
        /// </summary>
        [Browsable(false), DefaultValue(AutoSizeMode.GrowAndShrink)]
        public new AutoSizeMode AutoSizeMode
        {
            get { return base.AutoSizeMode; }
            set { base.AutoSizeMode = value; }
        }

        /// <summary>
        ///     This property is not relevant to the
        ///     <see cref="PlacesToolbar" /> class.
        /// </summary>
        [Browsable(false)]
        public override AutoValidate AutoValidate
        {
            get { return base.AutoValidate; }
            set { base.AutoValidate = value; }
        }

        /// <summary>
        ///     This property is not relevant to the
        ///     <see cref="PlacesToolbar" /> class.
        /// </summary>
        [Browsable(false)]
        public override Image BackgroundImage
        {
            get { return base.BackgroundImage; }
            set { base.BackgroundImage = value; }
        }

        /// <summary>
        ///     This property is not relevant to the
        ///     <see cref="PlacesToolbar" /> class.
        /// </summary>
        [Browsable(false)]
        public override ImageLayout BackgroundImageLayout
        {
            get { return base.BackgroundImageLayout; }
            set { base.BackgroundImageLayout = value; }
        }

        /// <summary>
        ///     This property is not relevant to the
        ///     <see cref="PlacesToolbar" /> class.
        /// </summary>
        [Browsable(false)]
        public new BorderStyle BorderStyle
        {
            get { return base.BorderStyle; }
            set { base.BorderStyle = value; }
        }

        /// <summary>
        ///     This property is not relevant to the
        ///     <see cref="PlacesToolbar" /> class.
        /// </summary>
        [Browsable(false)]
        public new bool CausesValidation
        {
            get { return base.CausesValidation; }
            set { base.CausesValidation = value; }
        }

        #endregion
    }
}