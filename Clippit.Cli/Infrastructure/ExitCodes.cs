namespace Clippit.Cli.Infrastructure;

internal static class ExitCodes
{
    public const int Success = 0;
    public const int InternalError = 1;
    public const int InvalidArguments = 2;
    public const int FileNotFound = 3;
    public const int InvalidFormat = 4;
    public const int OutputError = 5;
}

internal static class ErrorCodes
{
    public const string InternalError = "INTERNAL_ERROR";
    public const string InvalidArguments = "INVALID_ARGUMENTS";
    public const string FileNotFound = "FILE_NOT_FOUND";
    public const string InvalidFormat = "INVALID_FORMAT";
    public const string OutputError = "OUTPUT_ERROR";
}
