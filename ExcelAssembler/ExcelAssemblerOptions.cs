namespace ExcelAssembler;

public class ExcelAssemblerOptions
{
    public MissingXmlDataBehaviour MissingXmlDataBehaviour { get; set; } =
        MissingXmlDataBehaviour.ThrowException;
}