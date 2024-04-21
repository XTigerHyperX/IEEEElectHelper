using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEEEElectHelper
{
    internal class Writer
    {
        public static void writeskipped(string position , string ieeeID , string vote)
        {
            string name = Loader.returnname(ieeeID);
            string chapter = Functions.GetChapterName(position);
            string voteee = $"{position} , Skipped {ieeeID} ({name}) *** because is not a member in {chapter} , Vote was {vote}";
            Console.WriteLine($"Skipped {ieeeID} ({name}) *** because is not a member in {chapter}, Vote was {vote}  {position}");
            using (System.IO.StreamWriter file = new System.IO.StreamWriter("assets\\results\\skipped\\skipped.txt", true))
            {
                file.WriteLine($"Skipped vote details: {voteee}\n");
            }
        }
    }
}
