using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSOMExcercises.Constants
{
    static class Taxonomy
    {
        public const string TermGroupName = "CSOM Test";
        public const string TermSetName = "city-Nghiep";
        public enum TermsIndex
        {
            HoChiMinh,
            Stockholm
        }
        public static string GetTermName(TermsIndex term)
        {
            switch (term)
            {
                case TermsIndex.HoChiMinh:
                    return "Ho Chi Minh";
                case TermsIndex.Stockholm:
                    return "Stockholm";
                default:
                    return null;
            }
        }
    }
    namespace Columns
    {
        static class About
        {
            public const string Name = "About";
            public const string DisplayName = "About";
        }
        static class City
        {
            public const string Name = "City";
            public const string DisplayName = "City";
        }
        static class Author
        {
            public const string Name = "TestAuthor";
            public const string DisplayName = "Author";
        }
        static class Cities
        {
            public const string Name = "Cities";
            public const string DisplayName = "Cities";
        }
    }
    static class ContentType
    {
        public const string Name = "CSOM Test Content Type";
        public const string ParentContentTypeName = "Item";
    }
    static class List
    {
        public const string Name = "CSOM Test";
    }
    static class Document
    {
        public const string Name = "Document Test";
    }
}
