using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Upsideo.Agreagateur.Services
{
    public static class Helpers
    {
        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sequence"></param>
        /// <param name="nbrPerBlock"></param>
        /// <returns></returns>
        private static IEnumerable<ICollection<T>> SplitEnumerable<T>(this IEnumerable<T> sequence, int nbrPerBlock)
        {
            List<T> Group = new List<T>(nbrPerBlock);

            foreach (T value in sequence)
            {
                Group.Add(value);

                if (Group.Count == nbrPerBlock)
                {
                    yield return Group;
                    Group = new List<T>(nbrPerBlock);
                }
            }

            if (Group.Any()) yield return Group; // flush out any remaining
        }

        // now it's trivial; if you want to make smaller files, just foreach
        // over this and write out the lines in each block to a new file

        public static IEnumerable<ICollection<T>> SplitFile<T>(IEnumerable<T> sequence, int nbrPerBlock) where T :class
        {
            return sequence.SplitEnumerable(nbrPerBlock);
        }
    }
}