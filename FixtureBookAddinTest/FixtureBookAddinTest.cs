/*
 * Copyright 2014 XPFriend Community.
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
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XPFriend.Fixture;

namespace FixtureBookAddinTest
{
    [TestClass]
    public class FixtureBookAddinTest
    {
        [TestMethod]
        public void SQLServer__1000件データ()
        {
            FixtureBook.Expect(() => { });
        }

        [TestMethod]
        public void SQLServer__0件データ()
        {
            FixtureBook.Expect(() => { });
        }
        [TestMethod]
        public void SQLServer__1件データ()
        {
            FixtureBook.Expect(() => { });
        }
        [TestMethod]
        public void SQLServer__2件データ()
        {
            FixtureBook.Expect(() => { });
        }
        [TestMethod]
        public void SQLServer__3件データ()
        {
            FixtureBook.Expect(() => { });
        }

        [TestMethod]
        public void Oracle__1000件データ()
        {
            FixtureBook.Expect(() => { });
        }
    }
}
