using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections;
using System.Runtime.Serialization.Formatters.Binary;



namespace Test
{
     
   
    [Serializable]
    class DirData
    {
        public int Id;
        public string Name;
        public int ParentId;
    }

    //public static ArrayList list = new ArrayList();

    class Program
    {
        static public string filePath = @"C:\Users\HASEE\Desktop\test.dat";
        static void Main(string[] args)
        {

            //创建一个List，用于存储DirData的实例
            List<DirData> list = new List<DirData>();

            //chatgpt给的代码

            string fileName = @"C:\Users\HASEE\Desktop\NodeData.xlsx";

            // 根据文件扩展名创建工作簿对象
            IWorkbook workbook = null;
            if (Path.GetExtension(fileName) == ".xls")
            {
                workbook = new HSSFWorkbook(File.OpenRead(fileName));
            }
            else if (Path.GetExtension(fileName) == ".xlsx")
            {
                workbook = new XSSFWorkbook(File.OpenRead(fileName));
            }

            // 获取第一个工作表
            ISheet sheet = workbook.GetSheetAt(0);

            // 遍历行
            for (int i = 1; i <= sheet.LastRowNum; i++)
            {
                DirData treeData = new DirData();

                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    // 遍历列
                    for (int j = 0; j < row.LastCellNum; j++)
                    {
                        ICell cell = row.GetCell(j);


                        if (cell != null)
                        {
                            // 读取单元格中的值并输出
                            //Console.Write(cell.ToString());
                            //Console.WriteLine(cell.ToString());
                            //将数据添加到类中
                            if (j == 0)
                            {
                                treeData.Id = int.Parse(cell.ToString());
                            }
                            if (j == 1)
                            {
                                treeData.Name = cell.ToString();
                            }
                            if (j == 2)
                            {
                                treeData.ParentId = int.Parse(cell.ToString());
                            }

                        }
                    }
                }
                //打印数据
                //Console.WriteLine("ID = {0}     Name = {1}      ParentId = {2}",  treeData.Id, treeData.Name, treeData.ParentId);

                //将数据添加到list中
                list.Add(treeData);

            }

            //将List中的实例进行序列化并写入文件
            using (FileStream fs = new FileStream(filePath, FileMode.Create))
            {
                BinaryFormatter formatter = new BinaryFormatter();
                formatter.Serialize(fs, list);
                fs.Close();
            }


            //从文件中读取序列化后的数据并进行反序列化
            using (FileStream fs = new FileStream(filePath, FileMode.Open))
            {
                BinaryFormatter formatter = new BinaryFormatter();
                List<DirData> newList = (List<DirData>)formatter.Deserialize(fs);

                //输出反序列化后的数据
                foreach (var item in newList)
                {
                    //Console.WriteLine(item);
                    Console.WriteLine("输出反序列化后的数据");
                    Console.WriteLine("ID = {0}     Name = {1}      ParentId = {2}", item.Id, item.Name, item.ParentId);
                }
            }

            

            Console.ReadLine();

        }
    }
}
