/*
 * Copyright 2013 XPFriend Community.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
using System;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XPFriend.FixtureBook;

namespace FixtureBookAddinTest
{
    [TestClass]
    public class AddinTest
    {
        private CultureInfo defaultCulture;

        [TestInitialize]
        public void Setup()
        {
            defaultCulture = Thread.CurrentThread.CurrentUICulture;
        }

        [TestCleanup]
        public void Cleanup()
        {
            Thread.CurrentThread.CurrentUICulture = defaultCulture;
        }

        [TestMethod]
        public void 日本語でリボンUIのタブとグループとボタンが表示されること()
        {
            // setup
            Thread.CurrentThread.CurrentUICulture = CultureInfo.GetCultureInfo("ja-JP");
            RibbonUI ribbonUI = new RibbonUI();

            // expect
            XElement[] groups = ribbonUI.GetElements("group");
            Assert.AreEqual(2, groups.Count());
            Assert.AreEqual("追加", ribbonUI.GetLabel(groups[0]));
            Assert.AreEqual("設定", ribbonUI.GetLabel(groups[1]));

            XElement[] buttons = ribbonUI.GetElements("button");
            Assert.AreEqual(4, buttons.Count());
            Assert.AreEqual("シート", ribbonUI.GetLabel(buttons[0]));
            Assert.AreEqual("テストケース", ribbonUI.GetLabel(buttons[1]));
            Assert.AreEqual("テーブル", ribbonUI.GetLabel(buttons[2]));
            Assert.AreEqual("データベース", ribbonUI.GetLabel(buttons[3]));
        }

        [TestMethod]
        public void 英語でリボンUIのタブとグループとボタンが表示されること()
        {
            // setup
            Thread.CurrentThread.CurrentUICulture = CultureInfo.GetCultureInfo("en");
            RibbonUI ribbonUI = new RibbonUI();

            // expect
            XElement[] groups = ribbonUI.GetElements("group");
            Assert.AreEqual(2, groups.Count());
            Assert.AreEqual("Add", ribbonUI.GetLabel(groups[0]));
            Assert.AreEqual("Configuration", ribbonUI.GetLabel(groups[1]));

            XElement[] buttons = ribbonUI.GetElements("button");
            Assert.AreEqual(4, buttons.Count());
            Assert.AreEqual("New Sheet", ribbonUI.GetLabel(buttons[0]));
            Assert.AreEqual("Test case", ribbonUI.GetLabel(buttons[1]));
            Assert.AreEqual("Table", ribbonUI.GetLabel(buttons[2]));
            Assert.AreEqual("Database", ribbonUI.GetLabel(buttons[3]));
        }

        internal class RibbonUI
        {
            XElement doc;
            XmlNamespaceManager manager;
            public RibbonUI()
            {
                Addin addin = new Addin();
                string customUI = addin.GetCustomUI("XPFriend.FixtureBook.RibbonUI.xml");
                Console.WriteLine(customUI);
                doc = XElement.Parse(customUI);
                manager = new XmlNamespaceManager(new NameTable());
                manager.AddNamespace("ns", doc.GetDefaultNamespace().NamespaceName);
            }

            public XElement[] GetElements(string name)
            {
                return doc.XPathSelectElements("//ns:" + name, manager).ToArray();
            }

            public string GetAttribute(XElement element, string name)
            {
                return element.Attribute(XName.Get(name)).Value;
            }

            public string GetLabel(XElement element)
            {
                return element.Attribute(XName.Get("label")).Value;
            }
        }
    }
}
