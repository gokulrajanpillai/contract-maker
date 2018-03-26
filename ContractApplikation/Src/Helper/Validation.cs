using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ContractApplikation.Src.Helper
{
    public class Validation
    {
        struct RegexPatterns
        {
            public const string Email = @"^[A-Za-z._]+[@][A-Za-z._]+[.][A-Za-z._]+$";
            public const string PhoneNumber = @"^[+]*[0-9]{11,}$";
            public const string WholeNumber = @"^[0-9]+$";
            public const string ZipCode = @"^[0-9]{5}$";
            public const string Decimal = @"^[0-9]+[.]*[0-9]*$";
            public const string Characters = @"^[a-zA-Z]+$";
            public const string Name = @"^[A-Za-z.\s]*[A-Za-z]+$";
        }


        public static bool IsName(string name)
        {
            return new Regex(RegexPatterns.Name, RegexOptions.IgnoreCase).IsMatch(name);
        }

        public static bool IsValidEmail(string email)
        {
            return new Regex(RegexPatterns.Email, RegexOptions.IgnoreCase).IsMatch(email);
        }

        public static bool IsValidPhoneNumber(string number)
        {
            return Regex.Match(number, RegexPatterns.PhoneNumber).Success;
        }

        public static bool IsWholeNumber(string wholeNumber)
        {
            return Regex.Match(wholeNumber, RegexPatterns.WholeNumber).Success;
        }

        public static bool IsZipCode(string zipCode)
        {
            return Regex.Match(zipCode, RegexPatterns.ZipCode).Success;
        }

        public static bool IsDecimalNumber(string decimalNumber)
        {
            return Regex.Match(decimalNumber, RegexPatterns.Decimal).Success;
        }
    }
}
