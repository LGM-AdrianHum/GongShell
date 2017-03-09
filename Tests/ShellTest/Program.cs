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
using System.Collections.Generic;
using System.IO;
using GongSolutions.Shell;
using GongSolutions.Shell.Interop;
using NUnit.Framework;

namespace ShellTest
{
    public class Program
    {
        private static void Main(string[] args)
        {
            RunTests();
        }

        private static void RunTests()
        {
            var test = new ShellItemTest();

            test.SetUp();

            try
            {
                test.Names();
                test.EnumerateChildren();
                test.SpecialFolders();
                test.Attributes();
                test.Compare();
                test.Uri();
                test.Pidl();
            }
            finally
            {
                test.TearDown();
            }
        }
    }

    [TestFixture]
    public class ShellItemTest
    {
        [TestFixtureSetUp]
        public void SetUp()
        {
            Directory.CreateDirectory(Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "GongShellTestFolder"));
            Directory.CreateDirectory(Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "GongShellTestFolder\\Nested"));
        }

        [TestFixtureTearDown]
        public void TearDown()
        {
            Directory.Delete(Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "GongShellTestFolder"),
                true);
        }

        [Test]
        public void Attributes()
        {
            var cDrive = new ShellItem(@"C:\");
            var myComputer =
                new ShellItem(Environment.SpecialFolder.MyComputer);
            var file = new ShellItem(Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.System),
                "kernel32.dll"));

            Assert.IsTrue(cDrive.IsFolder);
            Assert.IsTrue(myComputer.IsFolder);
            Assert.IsFalse(file.IsFolder);
        }

        [Test]
        public void Compare()
        {
            var desktop1 = ShellItem.Desktop;
            var desktop2 = new ShellItem("shell:///Desktop");
            var item1 = new ShellItem(@"C:\Program Files");
            var item2 = item1.Parent[@"Program Files"];

            Assert.IsTrue(desktop1.Equals(desktop2));
            Assert.AreEqual(item1.FileSystemPath, item2.FileSystemPath);
            Assert.IsTrue(item1 == item2);
            Assert.IsFalse(item1 != item2);
        }

        [Test]
        public void EnumerateChildren()
        {
            var location = @"C:\Program Files";
            var item = new ShellItem(location);
            var children = new List<string>();

            children.AddRange(Directory.GetFiles(location));
            children.AddRange(Directory.GetDirectories(location));

            // The shell does not include desktop.ini in its enumeration.
            children.Remove(Path.Combine(location, "desktop.ini"));

            foreach (var child in item)
            {
                children.Remove(child.FileSystemPath);
            }

            Assert.IsEmpty(children);
        }

        [Test]
        public void Names()
        {
            var location = @"C:\Program Files";
            var displayName = @"Program Files";
            var item = new ShellItem(location);
            Assert.AreEqual(displayName, item.DisplayName);
            Assert.AreEqual(location, item.FileSystemPath);
            Assert.AreEqual(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                ShellItem.Desktop.FileSystemPath);
        }

        [Test]
        public void Pidl()
        {
            var myDocuments = new ShellItem(Environment.SpecialFolder.MyDocuments);
            var pidl = myDocuments.Pidl;
        }

        [Test]
        public void SpecialFolders()
        {
            var myComputerFolder =
                Environment.SpecialFolder.MyComputer;
            var myPicturesFolder =
                Environment.SpecialFolder.MyPictures;
            var myComputer = new ShellItem(myComputerFolder);
            var myComputer2 = new ShellItem(myComputer.ToUri());
            var myPictures = new ShellItem(myPicturesFolder);
            var myPictures2 = new ShellItem(myPictures.ToUri());

            Assert.IsTrue(myComputer.Equals(myComputer2));
            Assert.AreEqual(Environment.GetFolderPath(myPicturesFolder), myPictures.FileSystemPath);
            Assert.AreEqual(myPictures, myPictures2);

            try
            {
                var path = myComputer.FileSystemPath;
                Assert.Fail("FileSystemPath from virtual folder should throw exception");
            }
            catch (ArgumentException)
            {
            }
            catch (Exception)
            {
                Assert.Fail("FileSystemPath from virtual folder should throw ArgumentException");
            }
        }

        [Test]
        public void Uri()
        {
            var desktopExpected = "shell:///Desktop";
            var myComputerExpected = "shell:///MyComputerFolder";
            var myDocumentsExpected = "shell:///Personal";
            var desktop1 = ShellItem.Desktop;
            var desktop2 = new ShellItem(desktopExpected);
            var myComputer = new ShellItem(myComputerExpected);
            var myDocuments = new ShellItem(Environment.SpecialFolder.MyDocuments);
            var myDocumentsChild = myDocuments["GongShellTestFolder"];
            var myDocumentsChild2 = new ShellItem(myDocumentsExpected + "/GongShellTestFolder");
            var myDocumentsChild3 = new ShellItem(myDocumentsExpected + "/GongShellTestFolder/Nested");
            var cDrive = new ShellItem(@"C:\");
            var controlPanel = new ShellItem((Environment.SpecialFolder) CSIDL.CONTROLS);

            Assert.AreEqual(desktopExpected, desktop1.ToString());
            Assert.AreEqual(desktopExpected, desktop2.ToString());
            Assert.AreEqual(myComputerExpected, myComputer.ToString());
            Assert.AreEqual(myDocumentsExpected, myDocuments.ToString());
            Assert.AreEqual(myDocumentsExpected + "/GongShellTestFolder",
                myDocumentsChild.ToString());
            Assert.AreEqual("file:///C:/", cDrive.ToString());
            Assert.IsTrue(myDocumentsChild.Equals(myDocumentsChild2));
            Assert.AreEqual(myDocumentsExpected + "/GongShellTestFolder/Nested",
                myDocumentsChild3.ToString());
        }
    }
}