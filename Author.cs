//**********************************************************************************************
// File: Author.cs
//
// Purpose: Contains member variables, properties, and method definitions for class Author.
//
// Written By: Kevin Serina
// Compiler: Visual Studio 2019
//**********************************************************************************************
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.Globalization;
using System.Runtime.ExceptionServices;

namespace Publishing
{
    [DataContract]
    public class Author
    {
        #region Author Class Member Variables
        // Member variables - ALL PRIVATE
        private string first;
        private string last;
        private string background;
        #endregion

        #region Author Class Properties
        // Serializing members
        [DataMember(Name = "first")]
        public string First
        {
            get
            {
                return first;
            }
            set
            {
                first = value;
            }
        }

        [DataMember(Name = "last")]
        public string Last
        {
            get
            {
                return last;
            }
            set
            {
                last = value;
            }
        }

        [DataMember(Name = "background")]
        public string Background
        {
            get
            {
                return background;
            }
            set
            {
                background = value;
            }
        }
        #endregion

        #region Author Class Member Methods
        // Member Methods
        public Author()
        {
            first = "First";
            last = "Last";
            background = "Background: ";
        }
       
        public override string ToString()
        {
            return First + ", " + Last + ", " + Background;
        }
        #endregion
    }

    [DataContract]
    public class Book
    {
        #region Book Class Member Variables
        // Member Variables - ALL PRIVATE
        private string title;
        private double price;
        Author author = new Author();
        #endregion

        #region Book Class Properties
        // Properties
        [DataMember(Name = "title")]
        public string Title
        {
            get
            {
                return title;
            }
            set
            {
                title = value;
            }
        }
        [DataMember(Name = "price")]
        public double Price
        {
            get
            {
                return price;
            }
            set
            {
                price = value;
            }
        }

        [DataMember(Name = "author")]
        public Author Auth
        {
            get
            {
                return author;
            }
            set
            {
                author = value;
            }
        }
        #endregion

        #region Book Class Member Methods
        // Member Methods
        public Book()
        {
            title = "Title";
            price = 20;
        }
        #endregion

        public override string ToString()
        {
            return Title + ", " + author.First + ", " + author.Last + ", " + author.Background + ", " + Price;
        }
    }
}
