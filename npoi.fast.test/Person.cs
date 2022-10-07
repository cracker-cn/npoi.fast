namespace H.Npoi.Fast.Test
{
    public class Person
    {
        [Export("姓名")]
        public string? Name { get; set; }

        [Export("年龄")]
        public int Age { get; set; }

        [Export("爱好")]
        public string? Hobby { get; set; }
    }
}