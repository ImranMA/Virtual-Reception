using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VirtualReception
{
    public class VisitorInfo
    {
        public static string firstName { set; get; }
        public static string LastName { set; get; }

        public static bool UserContextTaken = false;

        public static bool PictureTaken = false;
    }


}
