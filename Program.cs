//**********************************************************************************************
// File: Main.cs
//
// Purpose: Program will contain a menu-driven program to manipulate and use the Author
//          and Book classes that I created. It will import the Publishing.dll solution.
//          When the program runs it will display a menu to the user and give them a 
//          chance to input a choice. An action will be taken depending on what choice 
//          the user makes. The menu actions should manipulate and use the appropriate 
//          class instance at the top of main.
//
// Written By: Kevin Serina
// Compiler: Visual Studio 2019
//**********************************************************************************************
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization.Json;
using System.Runtime.Serialization;
using System.Runtime.InteropServices;
using Publishing;
using System.IO;
using System.Security.Cryptography;
using System.ComponentModel.Design;
using System.Net;
using System.Diagnostics.Eventing.Reader;
using Microsoft.Office.Interop.Excel;

namespace PublishingConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            Author newAuthor = new Author();
            Book newBook = new Book();

            bool a = true;
            while (a)
            {
                Console.WriteLine("");
                Console.WriteLine("Publisher Menu");
                Console.WriteLine("--------------");
                Console.WriteLine("1 - Read Author from JSON file");
                Console.WriteLine("2 - Read Author from XML file");     
                Console.WriteLine("3 - Write Author to JSON file");     
                Console.WriteLine("4 - Write Author to XML file");      
                Console.WriteLine("5 - Write Author to Excel file");    
                Console.WriteLine("6 - Display Author data on screen"); 
                Console.WriteLine("7 - Read Book from JSON file");      
                Console.WriteLine("8 - Read Book from XML file");       
                Console.WriteLine("9 - Write Book to JSON file");       
                Console.WriteLine("10 - Write Book to XML file");       
                Console.WriteLine("11 - Write Book to Excel file");     
                Console.WriteLine("12 - Display Book data on screen");
                Console.WriteLine("13 - Exit");
                Console.Write("Enter choice: ");

                int choice = int.Parse(Console.ReadLine());
                switch (choice)
                {
                    case 1:
                        newAuthor = ReadAuthorJSON();
                        break;
                    case 2:
                        newAuthor = ReadAuthorXML();
                        break;
                    case 3:
                        WriteAuthorJSON(newAuthor);
                        break;
                    case 4:
                        WriteAuthorXML(newAuthor);
                        break;
                    case 5:
                        WriteAuthorExcel(newAuthor);
                        break;
                    case 6:
                        DisplayAuthor(newAuthor);
                        break;
                    case 7:
                        newBook = ReadBookJSON();
                        break;
                    case 8:
                        newBook = ReadBookXML();
                        break;
                    case 9:
                        WriteBookJSON(newBook);
                        break;
                    case 10:
                        WriteBookXML(newBook);
                        break;
                    case 11:
                        WriteBookExcel(newBook);
                        break;
                    case 12:
                        DisplayBook(newBook);
                        break;
                    case 13:
                        a = false;
                        break;
                }
            }
            
        }
        // Write Author to file Methods
        static void WriteAuthorJSON(Author author)
        {
            string fileName;
            
            Console.WriteLine("Enter file name: ");
            fileName = Console.ReadLine();

            Console.WriteLine("Please enter first name of the author");
            author.First = Console.ReadLine();

            Console.WriteLine("Please enter last name of the author");
            author.Last = Console.ReadLine();

            Console.WriteLine("Please enter background of the author");
            author.Background = Console.ReadLine();

            FileStream writer = new FileStream(fileName, FileMode.Create, FileAccess.Write);

            DataContractJsonSerializer json;
            json = new DataContractJsonSerializer(typeof(Author));

            json.WriteObject(writer, author);
            writer.Dispose();
        }

        static void WriteAuthorXML(Author author)
        {
            string fileName;

            Console.WriteLine("Enter file name: ");
            fileName = Console.ReadLine();

            Console.WriteLine("Please enter the first name of the author");
            author.First = Console.ReadLine();

            Console.WriteLine("Please enter the last name of the author");
            author.Last = Console.ReadLine();

            Console.WriteLine("Please enter the background of the author");
            author.Background = Console.ReadLine();

            FileStream writer = new FileStream(fileName, FileMode.Create, FileAccess.Write);

            DataContractSerializer xml;
            xml = new DataContractSerializer(typeof(Author));

            xml.WriteObject(writer, author);
            writer.Dispose();
        }

        static void WriteAuthorExcel(Author author)
        {
           
            string fileName = "";

            Console.WriteLine("Enter file name: ");
            fileName = Console.ReadLine();

            Console.WriteLine("Please enter the first name of the author");
            author.First = Console.ReadLine();

            Console.WriteLine("Please enter the last name of the author");
            author.Last = Console.ReadLine();

            Console.WriteLine("Please enter the background of the author");
            author.Background = Console.ReadLine();

            // Declare variables
            Application excelApp;
            Workbooks excelWorkBooks;
            _Workbook excelWorkBook;
            _Worksheet excelWorkSheet;

            // Start Excel and get Application object
            excelApp = new Application();
            excelApp.Visible = true;

            // Get a new workbook and worksheet
            excelWorkBooks = excelApp.Workbooks;
            excelWorkBook = (_Workbook)(excelWorkBooks.Add());
            excelWorkSheet = (_Worksheet)excelWorkBook.ActiveSheet;

            // Adds data to an Excel worksheet
            Range excelRange;
            excelRange = excelWorkSheet.Range["A1", "E1"];
            excelRange.Font.Bold = true;
            excelRange.Font.Underline = true;

            excelWorkSheet.Cells[1, 1] = "First";
            excelWorkSheet.Cells[1, 2] = "Last";
            excelWorkSheet.Cells[1, 3] = "Background";

            excelWorkSheet.Cells[2, 1] = author.First;
            excelWorkSheet.Cells[2, 2] = author.Last;
            excelWorkSheet.Cells[2, 3] = author.Background;

            // Save data
            excelWorkBook.SaveAs(fileName);
            excelWorkBook.Close();
            excelApp.Quit();

            // Make sure to release COM object
            if (excelRange != null)
                Marshal.ReleaseComObject(excelRange);
            if (excelWorkSheet != null) Marshal.ReleaseComObject(excelWorkSheet);
            if (excelWorkBook != null) Marshal.ReleaseComObject(excelWorkBook);
            if (excelWorkBooks != null) Marshal.ReleaseComObject(excelWorkBooks);
            if (excelApp != null) Marshal.ReleaseComObject(excelApp);
        }

        // Write Book to file Methods
        static void WriteBookJSON(Book book)
        {
            string fileName;

            Console.WriteLine("Enter file name: ");
            fileName = Console.ReadLine();

            Console.WriteLine("Please enter the title of the book");
            book.Title = Console.ReadLine();

            Console.WriteLine("Please enter first name of the author of the book");
            book.Auth.First = Console.ReadLine();

            Console.WriteLine("Please enter last name of the author");
            book.Auth.Last = Console.ReadLine();

            Console.WriteLine("Please enter background of the author");
            book.Auth.Background = Console.ReadLine();

            Console.WriteLine("Please enter the price of the book");
            book.Price = Convert.ToDouble(Console.ReadLine());

            FileStream writer = new FileStream(fileName, FileMode.Create, FileAccess.Write);

            DataContractJsonSerializer json;
            json = new DataContractJsonSerializer(typeof(Book));

            json.WriteObject(writer, book);
            writer.Dispose();
        }

        static void WriteBookXML(Book book)
        {
            string fileName;

            Console.WriteLine("Enter file name: ");
            fileName = Console.ReadLine();

            Console.WriteLine("Please enter the title of the book");
            book.Title = Console.ReadLine();

            Console.WriteLine("Please enter first name of the author");
            book.Auth.First = Console.ReadLine();

            Console.WriteLine("Please enter last name of the author");
            book.Auth.Last = Console.ReadLine();

            Console.WriteLine("Please enter the background of the author");
            book.Auth.Background = Console.ReadLine();

            Console.WriteLine("Please enter the price of the book");
            book.Price = Convert.ToDouble(Console.Read());

            FileStream writer = new FileStream(fileName, FileMode.Create, FileAccess.Write);

            DataContractSerializer xml;
            xml = new DataContractSerializer(typeof(Book));

            xml.WriteObject(writer, book);
            writer.Dispose();
        }        

        static void WriteBookExcel(Book book)
        {
            string fileName = "";

            Console.WriteLine("Enter file name: ");
            fileName = Console.ReadLine();

            Console.WriteLine("Please enter the title of the book");
            book.Title = Console.ReadLine();

            Console.WriteLine("Please enter the first name of the author");
            book.Auth.First = Console.ReadLine();

            Console.WriteLine("Please enter the last name of the author");
            book.Auth.Last = Console.ReadLine();

            Console.WriteLine("Please enter the background of the author");
            book.Auth.Background = Console.ReadLine();

            Console.WriteLine("Please enter the price of the book");
            book.Price = Convert.ToDouble(Console.Read());

            // Declare variables
            Application excelApp;
            Workbooks excelWorkBooks;
            _Workbook excelWorkBook;
            _Worksheet excelWorkSheet;

            // Start Excel and get Application object
            excelApp = new Application();
            excelApp.Visible = false;

            // Get a new workbook and worksheet
            excelWorkBooks = excelApp.Workbooks;
            excelWorkBook = (_Workbook)(excelWorkBooks.Add());
            excelWorkSheet = (_Worksheet)excelWorkBook.ActiveSheet;

            // Adds data to an Excel worksheet
            Range excelRange;
            excelRange = excelWorkSheet.Range["A1", "E1"];
            excelRange.Font.Bold = true;
            excelRange.Font.Underline = true;

            excelWorkSheet.Cells[1, 1] = "Title";
            excelWorkSheet.Cells[1, 2] = "First";
            excelWorkSheet.Cells[1, 3] = "Last";
            excelWorkSheet.Cells[1, 4] = "Background";
            excelWorkSheet.Cells[1, 5] = "Price";

            excelWorkSheet.Cells[2, 1] = book.Title;
            excelWorkSheet.Cells[2, 2] = book.Auth.First;
            excelWorkSheet.Cells[2, 3] = book.Auth.Last;
            excelWorkSheet.Cells[2, 4] = book.Auth.Background;
            excelWorkSheet.Cells[2, 5] = book.Price;

            // Save data
            excelWorkBook.SaveAs(fileName);
            excelWorkBook.Close();
            excelApp.Quit();

            // Make sure to release COM object
            if (excelRange != null)
                Marshal.ReleaseComObject(excelRange);
            if (excelWorkSheet != null) Marshal.ReleaseComObject(excelWorkSheet);
            if (excelWorkBook != null) Marshal.ReleaseComObject(excelWorkBook);
            if (excelWorkBooks != null) Marshal.ReleaseComObject(excelWorkBooks);
            if (excelApp != null) Marshal.ReleaseComObject(excelApp);
        }
        //***********************************************************************************************

        // Read Author from file Methods
        static Author ReadAuthorJSON()
        {
            string fileName;

            Console.WriteLine("Enter file name: ");
            fileName = Console.ReadLine();

            FileStream reader = new FileStream(fileName, FileMode.Open, FileAccess.Read);

            DataContractJsonSerializer inputSerializer;
            inputSerializer = new DataContractJsonSerializer(typeof(Author));

            
            Author author = (Author)inputSerializer.ReadObject(reader);
            reader.Dispose();
            return author;
        }

        static Author ReadAuthorXML()
        {
            string fileName;

            Console.WriteLine("Enter file name: ");
            fileName = Console.ReadLine();

            FileStream reader = new FileStream(fileName, FileMode.Open, FileAccess.Read);

            DataContractSerializer inputSerializer;
            inputSerializer = new DataContractSerializer(typeof(Author));

            Author author = (Author)inputSerializer.ReadObject(reader);
            reader.Dispose();
            return author;
        }

        // Read Book from file Methods

        static Book ReadBookJSON()
        {
            string fileName;

            Console.WriteLine("Enter file name: ");
            fileName = Console.ReadLine();

            FileStream reader = new FileStream(fileName, FileMode.Open, FileAccess.Read);

            DataContractJsonSerializer inputSerializer;
            inputSerializer = new DataContractJsonSerializer(typeof(Book));

            Book book = (Book)inputSerializer.ReadObject(reader);
            reader.Dispose();
            return book;
        }

        static Book ReadBookXML()
        {
            string fileName;
            Console.WriteLine("Enter file name: ");
            fileName = Console.ReadLine();

            FileStream reader = new FileStream(fileName, FileMode.Open, FileAccess.Read);

            DataContractSerializer inputSerializer;
            inputSerializer = new DataContractSerializer(typeof(Book));

            Book book = (Book)inputSerializer.ReadObject(reader);
            reader.Dispose();
            return book;           
        }

        // Display Author Data to screen
        static void DisplayAuthor(Author author)
        {
            Console.WriteLine(author.ToString());   
        }

        // Display Book Data on screen
        static void DisplayBook(Book book)
        {
            Console.WriteLine(book.ToString());
        }
    }
}

