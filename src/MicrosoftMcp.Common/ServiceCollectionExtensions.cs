using Azure.Core;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;
using Microsoft.Graph;

namespace MicrosoftMcp.Common;

public static class ServiceCollectionExtensions
{
    /// <summary>Registers GraphAuthOptions + TokenCredential + GraphServiceClient (singleton per local stdio process).</summary>
    public static IServiceCollection AddGraphCommon(this IServiceCollection services, IConfiguration configuration)
    {
        services.Configure<GraphAuthOptions>(configuration.GetSection(GraphAuthOptions.SectionName));
        services.AddSingleton<IValidateOptions<GraphAuthOptions>, GraphAuthOptionsValidator>();
        services.AddSingleton<ITokenCredentialProvider, TokenCredentialFactory>();

        services.AddSingleton<TokenCredential>(sp =>
        {
            var options = sp.GetRequiredService<IOptions<GraphAuthOptions>>().Value;
            return sp.GetRequiredService<ITokenCredentialProvider>().GetCredential(options);
        });

        services.AddSingleton<GraphServiceClient>(sp =>
        {
            var options = sp.GetRequiredService<IOptions<GraphAuthOptions>>().Value;
            var credential = sp.GetRequiredService<TokenCredential>();
            string[] scopes = options.AuthMode == AuthMode.AppOnly
                ? ["https://graph.microsoft.com/.default"]
                : options.DelegatedScopes;
            return new GraphServiceClient(credential, scopes);
        });

        return services;
    }
}
