namespace ZOfit.Installer;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        bool silent = args.Any(a => string.Equals(a, "--silent", StringComparison.OrdinalIgnoreCase));
        bool uninstall = args.Any(a => string.Equals(a, "--uninstall", StringComparison.OrdinalIgnoreCase));

        if (silent)
        {
            try
            {
                var svc = new InstallService();
                if (uninstall)
                    svc.Uninstall(msg => Console.WriteLine(msg));
                else
                    svc.Install(msg => Console.WriteLine(msg));
                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("ERROR: " + ex.Message);
                return 1;
            }
        }

        ApplicationConfiguration.Initialize();
        Application.Run(new MainForm());
        return 0;
    }
}
