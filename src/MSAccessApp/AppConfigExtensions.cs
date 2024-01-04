using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Configuration;

namespace MSAccessApp
{
    public static class AppConfigExtensions
    {
        public static IConfigurationRoot LoadConfig()
        {
            var config = new ConfigurationBuilder();
            config.AddJsonFile("appsettings.json");

            if (File.Exists("appSettings.Development.json"))
                config.AddJsonFile("appSettings.Development.json");

            return config.Build();
        }

        public static IServiceCollection ConfigureServices(this IServiceCollection services, IConfigurationRoot config)
        {
            services.AddSingleton(config);
            services.AddLogging(loggerBuilder =>
            {
                loggerBuilder.AddConsole();
            });

            return services;
        }

    }
}
