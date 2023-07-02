using FileRw;
using System.Collections.Generic;

namespace Test
{
    public class DataOperate
    {
        ////list合并同类项目
        //public static List<Type> Statisics()
        //{
        //    List<Type> list = new List<Type>();

        //    return list;
        //}
        //list合并同类项目
        public static List<CodeData> Statisics(List<CodeData> list)
        {

            //先给list排个序
            list.Sort((p1, p2) => p1.Id.CompareTo(p2.Id));
            //合并同类项
            for (int i = 1; i < list.Count; i++)
            {
                while(list[i - 1].Id == list[i].Id)
                {
                    list[i - 1].num = list[i - 1].num + list[i].num;
                    list.RemoveAt(i);
                }
                //for(int j = i + 1; j  < i; j++)
                //{
                //    if (list[i] == list[j])
                //    {

                //    }
                //}    
            }


            return list;
        }
    }
}
