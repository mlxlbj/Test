using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;



//namespace FileOperate
namespace FileRw
{

    [Serializable]

    //数据类
    class CodeData
    {
        public string Id = " ";                 //编码
        public string name = " ";               //名称
        public string specifications = " ";     //规格
        public string brand = " ";              //品牌
        public string material = " ";           //材质
        public double num;                     //数量
        public string unit = "PCS";             //单位
        public string notes = " ";              //备注
    }


    class ExcelOperate
    {
        //无参默认构造函数
        public ExcelOperate()
        {

            //创建一个List，用于存储DirData的实例
            List<CodeData> list = new List<CodeData>();


            //读取Excel中的数据
            string fileName = @"C:\Users\admin\Desktop\B-BOM数据汇总.xlsx";

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

            // 遍历行， 从第六行开始读取数据
            for (int i = 5; i <= sheet.LastRowNum; i++)
            {
                //临时存放数据
                CodeData tempData = new CodeData();

                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    // 遍历列  从第二列开始读取数据
                    //for (int j = 0; j < row.LastCellNum; j++)
                    for (int j = 1; j < 9; j++)
                    {
                        ICell cell = row.GetCell(j);


                        if (cell != null)
                        {
                            //读取单元格中的值，并将数据添加到类中
                            if (j == 1)
                            {
                                tempData.Id = cell.ToString();
                            }
                            if (j == 2)
                            {
                                tempData.name = cell.ToString();
                            }
                            if (j == 3)
                            {
                                tempData.specifications = cell.ToString();
                            }
                            if (j == 4)
                            {
                                tempData.brand = cell.ToString();
                            }
                            if (j == 5)
                            {
                                tempData.material = cell.ToString();
                            }
                            if (j == 6)
                            {
                                //执行前需要判断是否为空 or 或是否为小数
                                //tempData.num = int.Parse(cell.ToString());
                                tempData.num = double.Parse(cell.ToString());
                            }
                            if (j == 7)
                            {
                                tempData.unit = cell.ToString();
                            }
                            if (j == 8)
                            {
                                tempData.notes = cell.ToString();
                            }

                        }
                    }
                }
                //打印数据
                //Console.WriteLine("ID = {0}     Name = {1}      ParentId = {2}",  treeData.Id, treeData.Name, treeData.ParentId);
                Console.WriteLine("编码 = {0}     名称 = {1}            规格 = {2}          品牌 = {3}    材质 = {4}    数量 = {5}    单位 = {6}", tempData.Id, tempData.name, tempData.specifications, tempData.brand, tempData.material, tempData.num, tempData.unit);

                //将数据添加到list中
                //list.Add(tempData);
                //list.Sort();
                //foreach (tempData in list)
                //{
                //    temp
                //}
            }
        }
    }
    //public class A
    //{
    //    public A()
    //    {
    //        //将List中的实例进行序列化并写入文件
    //        using (FileStream fs = new FileStream(filePath, FileMode.Create))
    //        {
    //            BinaryFormatter formatter = new BinaryFormatter();
    //            formatter.Serialize(fs, list);
    //            fs.Close();
    //        }
    //    }
        
    //    public Myread()
    //    {
    //        //从文件中读取序列化后的数据并进行反序列化
    //        using (FileStream fs = new FileStream(filePath, FileMode.Open))
    //        {
    //            BinaryFormatter formatter = new BinaryFormatter();
    //            List<DirData> newList = (List<DirData>)formatter.Deserialize(fs);

    //            //输出反序列化后的数据
    //            foreach (var item in newList)
    //            {
    //                //Console.WriteLine(item);
    //                Console.WriteLine("输出反序列化后的数据");
    //                Console.WriteLine("ID = {0}     Name = {1}      ParentId = {2}", item.Id, item.Name, item.ParentId);
    //            }
    //        }
    //    }       
    //}
    
 }  




