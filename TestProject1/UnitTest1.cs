namespace TestProject1
{
    public class UnitTest1
    {
        [Fact]
        public void Test1()
        {
            var p = GeradorRelatorioPDF.Program.DesserializarPessoas();
            Assert.True(p.Count > 0);
        }
    }
}