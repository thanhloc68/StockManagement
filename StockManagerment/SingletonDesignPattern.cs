using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StockManagerment
{
    internal class Singleton
    {
        private static Singleton instance = null;
        private static object syncRoot = new object();

        private Singleton() { }
        public static Singleton GetInstance
        {
            get
            {
                lock (syncRoot)
                {
                    if (instance == null)
                    {
                        instance = new Singleton();
                    }
                }
                return instance;
            }
        }
        public void Getdb()
        {
            StockDataContext dbcontext = new StockDataContext();

        }
    }
}
