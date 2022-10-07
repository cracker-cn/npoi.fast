/*
 * @Author: HUA 
 * @Date: 2022-10-07 11:13:49 
 * @Last Modified by:   HUA 
 * @Last Modified time: 2022-10-07 11:13:49 
 */
using H.Npoi.Fast;
using H.Npoi.Fast.Test;
using System.Data;

namespace H.Npoi.Fast.test
{
    internal class Program
    {
        static void Main(string[] args)
        {
            BaseExport<Person> baseExport = new BaseExport<Person>();
            Person person = new Person()
            {
                Name="小明",
                Age=19,
                Hobby="篮球"
            };
            baseExport.Do(person);
        }
    }
}