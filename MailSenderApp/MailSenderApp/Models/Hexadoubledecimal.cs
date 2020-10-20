namespace MailSenderApp.Models
{
    #region Namespace.
    using System;
    using System.Collections.Generic;
    using System.Text.RegularExpressions;
    #endregion

    public class Hexadoubledecimal
    {
        public Hexadoubledecimal(int value)
        {
            if (value > 0)
            {
                this.DecValue = value;
                this.HDDValue = this.DecToHDD(value);
            }
        }

        public Hexadoubledecimal(string value)
        {
            if (this.CheckHDDFormat(value))
            {
                this.DecValue = this.HDDToDec(value);
                this.HDDValue = value;
            }
        }

        public int DecValue { get; private set; }
        public string HDDValue { get; private set; }

        public int HDDToDec(string value)
        {
            if (!this.CheckHDDFormat(value))
            {
                return 0;
            }

            int res = 0;
            for (int i = 0; i < value.Length; i++)
            {
                res += Convert.ToInt32(Mappings[value[value.Length - 1 - i]] * Math.Pow(26, i));
            }

            return res;
        }

        public string DecToHDD(int value)
        {
            if (value <= 0)
            {
                return string.Empty;
            }

            string res = string.Empty;
            int q = value;

            while (true)
            {
                int c = q % 26;
                q = q / 26;
                if (c == 0 && q != 0)
                {
                    res += 'Z';
                    q--;
                }
                else
                {
                    res += InverseMappings[c - 1];
                }
                
                if (q <= 26 && q > 0)
                {
                    res += InverseMappings[q - 1];

                    break;
                }

                if (q == 0)
                {
                    break;
                }
            }

            char[] arr = res.ToCharArray();
            Array.Reverse(arr);

            return new String(arr);
        }

        private bool CheckHDDFormat(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return false;
            }

            string pattern = @"^[A-Z]+$";
            foreach (var match in Regex.Matches(value, pattern))
            {
                return match.ToString() == value;
            }

            return false;
        }

        private static Dictionary<char, int> Mappings = new Dictionary<char, int>()
        {
            { 'A', 1 }, { 'B', 2 }, { 'C', 3 }, { 'D', 4 }, { 'E', 5 },
            { 'F', 6 }, { 'G', 7 }, { 'H', 8 }, { 'I', 9 }, { 'J', 10 },
            { 'K', 11 }, { 'L', 12 }, { 'M', 13 }, { 'N', 14 }, { 'O', 15 },
            { 'P', 16 }, { 'Q', 17 }, { 'R', 18 }, { 'S', 19 }, { 'T', 20 },
            { 'U', 21 }, { 'V', 22 }, { 'W', 23 }, { 'X', 24 }, { 'Y', 25 },
            { 'Z', 26 }
        };

        private static char[] InverseMappings = new char[26] 
        {
            'A', 'B', 'C', 'D', 'E',
            'F', 'G', 'H', 'I', 'J',
            'K', 'L', 'M', 'N', 'O',
            'P', 'Q', 'R', 'S', 'T',
            'U', 'V', 'W', 'X', 'Y',
            'Z'
        };
    }
}
