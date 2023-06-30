using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections;
using System.Runtime.Serialization.Formatters.Binary;
using FileRw;



namespace Test
{
     
   
    [Serializable]
    //class DirData
    //{
    //    public int Id;
    //    public string Name;
    //    public int ParentId;
    //}

    //public static ArrayList list = new ArrayList();

    class Program
    {

        static void Main(string[] args)
        {
            //初始化excel对象
            ExcelOperate myexcel = new ExcelOperate();

            Console.ReadLine();

        }
    }
}
