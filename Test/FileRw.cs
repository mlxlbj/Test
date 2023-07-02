using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using Test;



//namespace FileOperate
namespace FileRw
{

    [Serializable]

    //数据类
    public class CodeData
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


    public class ExcelOperate
    {
        public static List<CodeData> dlist = new List<CodeData>();


        //从excel中读取数据
        #region ReadExeclA
        public static void ReadExcelA(string fileName)
        {

            //创建一个List，用于存储DirData的实例
            List<CodeData> list = new List<CodeData>();

            //读取Excel中的数据
            //string fileName = @"C:\Users\admin\Desktop\B-BOM数据汇总.xlsx";

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
                    //将数据添加到list中
                    
                    dlist.Add(tempData);
                }
                //打印数据
                //Console.WriteLine("ID = {0}     Name = {1}      ParentId = {2}",  treeData.Id, treeData.Name, treeData.ParentId);
                //Console.WriteLine("编码 = {0}     名称 = {1}            规格 = {2}          品牌 = {3}    材质 = {4}    数量 = {5}    单位 = {6}", tempData.Id, tempData.name, tempData.specifications, tempData.brand, tempData.material, tempData.num, tempData.unit);

                
                //list.Sort();
                //foreach (tempData in list)
                //{
                //    temp
                //}
            }
            //调用汇总函数
            dlist = DataOperate.Statisics(dlist);
        }
        #endregion

        #region riteExcelA
        //将数据写入excel中
        public static void WriteExcelA()
        {
            IWorkbook wkbook = new HSSFWorkbook();
            ISheet sheet = wkbook.CreateSheet("sheetName");
            sheet.SetColumnWidth(0, 15 * 320);
            sheet.SetColumnWidth(1, 15 * 420);
            sheet.SetColumnWidth(2, 15 * 1200);
            sheet.SetColumnWidth(3, 15 * 260);
            sheet.SetColumnWidth(4, 15 * 260);
            sheet.SetColumnWidth(5, 15 * 185);
            sheet.SetColumnWidth(6, 15 * 120);
            bool wt = true;
            CodeData tempdata;
            Console.WriteLine(dlist.Count);

            for (int i = 0; i < dlist.Count + 1; i++)
            {
                IRow row = sheet.CreateRow(i);
                if (wt)
                {
                    row.CreateCell(0).SetCellValue("编码");
                    row.CreateCell(1).SetCellValue("名称");
                    row.CreateCell(2).SetCellValue("规格");
                    row.CreateCell(3).SetCellValue("品牌");
                    row.CreateCell(4).SetCellValue("材质");
                    row.CreateCell(5).SetCellValue("数量");
                    row.CreateCell(6).SetCellValue("单位");
                    row.CreateCell(7).SetCellValue("备注");
                    wt = false;
                }
                else
                {
                    //写入数据
                    //row.CreateCell(0).SetCellValue("标题");
                    tempdata = (CodeData)dlist[i - 1];
                    row.CreateCell(0).SetCellValue((string)tempdata.Id);
                    row.CreateCell(1).SetCellValue((string)tempdata.name);
                    row.CreateCell(2).SetCellValue((string)tempdata.specifications);
                    row.CreateCell(3).SetCellValue((string)tempdata.brand);
                    row.CreateCell(4).SetCellValue((string)tempdata.material);
                    row.CreateCell(5).SetCellValue((int)tempdata.num);
                    row.CreateCell(6).SetCellValue((string)tempdata.unit);
                    row.CreateCell(7).SetCellValue((string)tempdata.notes);

                }
            }
            //将信息写入文件
            using (FileStream fsWrite = File.OpenWrite(@"../输出BOM汇总.xlsx"))
            {
                wkbook.Write(fsWrite);
            }
        }
        #endregion
    }
    //public class BinaryOperate
    //{
    //    public static  void WriteBinaryA(string filePath, List<Type> list)
    //    {
    //        //将List中的实例进行序列化并写入文件
    //        using (FileStream fs = new FileStream(filePath, FileMode.Create))
    //        {
    //            BinaryFormatter formatter = new BinaryFormatter();
    //            formatter.Serialize(fs, list);
    //            fs.Close();
    //        }
    //    }

    //    public void Myread<T>(string filePath)
    //    {
    //        //从文件中读取序列化后的数据并进行反序列化
    //        using (FileStream fs = new FileStream(filePath, FileMode.Open))
    //        {
    //            BinaryFormatter formatter = new BinaryFormatter();
    //            //List<DirData> newList = (List<DirData>)formatter.Deserialize(fs);
    //            List<T> newList = (List<T>)formatter.Deserialize(fs);

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




