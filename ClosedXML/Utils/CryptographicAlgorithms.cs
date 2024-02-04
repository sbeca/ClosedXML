#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using static ClosedXML.Excel.XLProtectionAlgorithm;

namespace ClosedXML.Utils
{
    internal static class CryptographicAlgorithms
    {
#if NET6_0_OR_GREATER
        private static readonly byte[] nonZeroBytes = new byte[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123, 124, 125, 126, 127, 128, 129, 130, 131, 132, 133, 134, 135, 136, 137, 138, 139, 140, 141, 142, 143, 144, 145, 146, 147, 148, 149, 150, 151, 152, 153, 154, 155, 156, 157, 158, 159, 160, 161, 162, 163, 164, 165, 166, 167, 168, 169, 170, 171, 172, 173, 174, 175, 176, 177, 178, 179, 180, 181, 182, 183, 184, 185, 186, 187, 188, 189, 190, 191, 192, 193, 194, 195, 196, 197, 198, 199, 200, 201, 202, 203, 204, 205, 206, 207, 208, 209, 210, 211, 212, 213, 214, 215, 216, 217, 218, 219, 220, 221, 222, 223, 224, 225, 226, 227, 228, 229, 230, 231, 232, 233, 234, 235, 236, 237, 238, 239, 240, 241, 242, 243, 244, 245, 246, 247, 248, 249, 250, 251, 252, 253, 254, 255 };
#endif

        public static string Base64Decode(string base64EncodedData)
        {
            var base64EncodedBytes = Convert.FromBase64String(base64EncodedData);
            return Encoding.UTF8.GetString(base64EncodedBytes);
        }

        public static string Base64Encode(string plainText)
        {
            var plainTextBytes = Encoding.UTF8.GetBytes(plainText);
            return Convert.ToBase64String(plainTextBytes);
        }

        public static String GenerateNewSalt(Algorithm algorithm)
        {
            if (RequiresSalt(algorithm))
                return GetSalt();
            else
                return String.Empty;
        }

        public static String GetPasswordHash(Algorithm algorithm, String password, String salt = "", UInt32 spinCount = 0)
        {
            if (password == null)
                throw new ArgumentNullException(nameof(password));

            if (salt == null)
                throw new ArgumentNullException(nameof(salt));

            if (password.Length == 0) return "";

            switch (algorithm)
            {
                case Algorithm.SimpleHash:
                    return GetDefaultPasswordHash(password);

                case Algorithm.SHA512:
                    return GetSha512PasswordHash(password, salt, spinCount);

                default:
                    return string.Empty;
            }
        }

        public static string GetSalt(int length = 32)
        {
#if NET6_0_OR_GREATER
            var salt = RandomNumberGenerator.GetItems<byte>(nonZeroBytes, length);
            return Convert.ToBase64String(salt);
#else
            using (var random = new RNGCryptoServiceProvider())
            {
                var salt = new byte[length];
                random.GetNonZeroBytes(salt);
                return Convert.ToBase64String(salt);
            }
#endif
        }

        public static Boolean RequiresSalt(Algorithm algorithm)
        {
            switch (algorithm)
            {
                case Algorithm.SimpleHash:
                    return false;

                case Algorithm.SHA512:
                    return true;

                default:
                    return false;
            }
        }

        private static String GetDefaultPasswordHash(String password)
        {
            if (password == null)
                throw new ArgumentNullException(nameof(password));

            // http://kohei.us/2008/01/18/excel-sheet-protection-password-hash/
            // http://sc.openoffice.org/excelfileformat.pdf - 4.18.4
            // http://web.archive.org/web/20080906232341/http://blogs.infosupport.com/wouterv/archive/2006/11/21/Hashing-password-for-use-in-SpreadsheetML.aspx
            byte[] passwordCharacters = Encoding.ASCII.GetBytes(password);
            int hash = 0;
            if (passwordCharacters.Length > 0)
            {
                int charIndex = passwordCharacters.Length;

                while (charIndex-- > 0)
                {
                    hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
                    hash ^= passwordCharacters[charIndex];
                }
                // Main difference from spec, also hash with charcount
                hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
                hash ^= passwordCharacters.Length;
                hash ^= (0x8000 | ('N' << 8) | 'K');
            }

            return Convert.ToString(hash, 16).ToUpperInvariant();
        }

        private static String GetSha512PasswordHash(String password, String salt, UInt32 spinCount)
        {
            if (password == null)
                throw new ArgumentNullException(nameof(password));

            if (salt == null)
                throw new ArgumentNullException(nameof(salt));

            var saltBytes = Convert.FromBase64String(salt);
            var passwordBytes = Encoding.Unicode.GetBytes(password);
            var bytes = saltBytes.Concat(passwordBytes).ToArray();

            byte[] hashedBytes;
#if NET6_0_OR_GREATER
            using (var hash = SHA512.Create())
#else
            using (var hash = new SHA512Managed())
#endif
            {
                hashedBytes = hash.ComputeHash(bytes);

                bytes = new byte[hashedBytes.Length + sizeof(uint)];
                for (uint i = 0; i < spinCount; i++)
                {
                    var le = BitConverter.GetBytes(i);
                    Array.Copy(hashedBytes, bytes, hashedBytes.Length);
                    Array.Copy(le, 0, bytes, hashedBytes.Length, le.Length);
                    hashedBytes = hash.ComputeHash(bytes);
                }
            }

            return Convert.ToBase64String(hashedBytes);
        }
    }
}
