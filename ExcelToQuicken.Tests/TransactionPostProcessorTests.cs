namespace BanortePDFToExcel.Tests;

public sealed class TransactionPostProcessorTests
{
    private readonly TransactionPostProcessor _processor = new();

    [Theory]
    [InlineData("TRASPASO, DE LA CUENTA: 1171346382", "[Banorte Negocio]")]
    [InlineData("CONCENTRACION COMISION FEDERAL DE ELECTRICIDAD", "Home:Bills & Utilities:Utilities:Electrical")]
    [InlineData("COSTCO GAS", "Auto & Transport:Gas & Fuel")]
    public void DetermineCategory_UsesExistingCodeRules(string description, string expectedCategory)
    {
        Assert.Equal(expectedCategory, _processor.DetermineCategory(description));
    }

    [Fact]
    public void DetermineCategory_ReturnsEmptyForUnknownMerchant()
    {
        Assert.Equal(string.Empty, _processor.DetermineCategory("Unknown merchant"));
    }

    [Fact]
    public void FormatTransaction_PrioritizesBeneficiaryAndConceptForQuicken()
    {
        var result = _processor.FormatTransaction(
            "COMPRA   ORDEN DE PAGO SPEI 0260403 =REFERENCIA CTA/CLABE: 4152314364776081 , BXI SPEI BCO:012 BENEF:LORENA YOSUHA   RAMOS SANCHE (DATO NO VERIF POR ESTA INST), Homeschool teacher CVE RASTREO: 38432P05202604035153553752 RFC: ND IVA: . 0 BBVA MEXICO HORA LIQ: 12:15:19");

        Assert.True(result.Description.Length <= 64);
        Assert.Contains("COMPRA ORDEN DE PAGO SPEI", result.Description);
        Assert.Contains("LORENA", result.Description);
        Assert.Contains("BENEF: LORENA YOSUHA RAMOS SANCHE", result.Memo);
        Assert.Contains("Concepto: Homeschool teacher", result.Memo);
        Assert.DoesNotContain("CVE RASTREO", result.Description);
    }

    [Fact]
    public void FormatTransaction_LimitsMemoToQuickenMaximum()
    {
        var result = _processor.FormatTransaction(
            "COMPRA ORDEN DE PAGO SPEI 0260530 Guillermo israe Semana de",
            new string('x', 320));

        Assert.True(result.Memo.Length <= 255);
        Assert.EndsWith("…", result.Memo);
    }
}
