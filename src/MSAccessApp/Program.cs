// See https://aka.ms/new-console-template for more information

using Microsoft.Extensions.DependencyInjection;
using MSAccessApp;

var cfg = AppConfigExtensions.LoadConfig();
var provider = new ServiceCollection()
    .ConfigureServices(cfg)
    .BuildServiceProvider();

var t = new Examples(provider);
t.Run();
