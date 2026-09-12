using Azure.Core;

namespace MicrosoftMcp.Common;

/// <summary>
/// Creates the <see cref="TokenCredential"/> for Graph. Abstracted so a future
/// central HTTP host can plug in an On-Behalf-Of provider without touching Outlook code.
/// </summary>
public interface ITokenCredentialProvider
{
    TokenCredential GetCredential(GraphAuthOptions options);
}
