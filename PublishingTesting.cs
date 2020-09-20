using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Publishing;

namespace UnitTestProject
{
    
    [TestClass]
    public class AuthorTests
    {
        
        [TestMethod]
        public void TestFirst()
        {
            // Arrange
            Author author = new Author();
            string name = "John";
            author.First = "John";

            Assert.IsTrue(author.First == name, "TestFirst method has failed.");
        }

        [TestMethod]
        public void TestLast()
        {
            Author author = new Author();
            author.Last = "Smith";

            Assert.IsTrue(author.Last == "Smith", "TestLast method has failed");
        }

        [TestMethod]
        public void TestBackground()
        {
            Author author = new Author();
            author.Background = "No background";
            Assert.IsTrue(author.Background == "No background", "TestBackground method has failed.");
        }
    }

    [TestClass]
    public class BookTests
    {
        [TestMethod]
        public void TestTitle()
        {
            Book book = new Book();
            book.Title = "Lord of the Rings";
            Assert.IsTrue(book.Title == "Lord of the Rings", "TestBackground method has failed.");
        }

        [TestMethod]
        public void TestAuthor()
        {
            Book book = new Book();

            string first = "JK";
            string last = "Rowling";
            string backGround = "British writer and philanthropist.";

            book.Auth.First = "JK";
            book.Auth.Last = "Rowling";
            book.Auth.Background = "British writer and philanthropist.";

            Assert.IsTrue(book.Auth.First == first, "TestAuthor method has failed.");
            Assert.IsTrue(book.Auth.Last == last, "TestAuthor method has failed.");
            Assert.IsTrue(book.Auth.Background == backGround, "TestAuthor method has failed.");
        }

        [TestMethod]
        public void TestPrice()
        {
            Book book = new Book();
            book.Price = 25;
            double actual = 25;
            Assert.AreEqual(book.Price, actual);
        }
    }
}
