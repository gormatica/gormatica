using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace gormatica
{
    public partial class gormatica : Form
    {
        // Rutas
        private readonly string Carpeta =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Gormaz Informática");
        private readonly string CarpetaDatos =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Gormaz Informática");

        private readonly string TvExeName = "TeamViewerQS-idcs946vrj.exe";
        private string TeamViewerPath { get { return Path.Combine(Carpeta, TvExeName); } }

        // Red
        private static readonly HttpClient http;

        // Cancelación
        private readonly CancellationTokenSource _cts = new CancellationTokenSource();

        static gormatica()
        {
            // TLS modernos en .NET Framework
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;

            // Descompresión compatible (GZip/Deflate). Nada de DecompressionMethods.All aquí.
            var handler = new HttpClientHandler
            {
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
            };

            http = new HttpClient(handler)
            {
                Timeout = TimeSpan.FromSeconds(20)
            };
        }

        public gormatica(string[] args)
        {
            InitializeComponent();

            if (progreso != null)
            {
                progreso.Minimum = 0;
                progreso.Maximum = 100;
                progreso.Value = 0;
                progreso.Style = ProgressBarStyle.Blocks;
                progreso.Visible = false;
            }
        }

        // ================== ACTUALIZACIÓN ==================

        private async Task CheckUpdateAsync()
        {
            SetStatus("Buscando actualizaciones…");
            try
            {
                string latestText = await GetLatestVersionAsync(_cts.Token);
                Version latest, current;

                if (!string.IsNullOrWhiteSpace(latestText)
                    && Version.TryParse(latestText.Trim(), out latest)
                    && Version.TryParse(Application.ProductVersion, out current)
                    && latest > current)
                {
                    SafeUI(delegate
                    {
                        label_quedeseas.Visible = false;
                        button_teamviewer.Visible = false;
                        version.Text = "Versión " + Application.ProductVersion + " · ¡versión " + latest + " disponible!";
                    });

                    for (int i = 3; i >= 0; i--)
                    {
                        SetStatus("Actualizando en " + i + " s …");
                        await Task.Delay(1000);
                    }

                    await LaunchUpdaterAsync(_cts.Token);
                    return;
                }
            }
            catch (OperationCanceledException) { }
            catch (Exception ex)
            {
                SetStatus("Error al comprobar actualización: " + ex.Message);
            }

            SetStatus(string.Empty);
        }

        private async Task<string> GetLatestVersionAsync(CancellationToken ct)
        {
            const string url = "https://www.gormatica.com/servicios/soporte-remoto/version.txt";
            using (var resp = await http.GetAsync(url, ct))
            {
                resp.EnsureSuccessStatusCode();
                // En .NET Framework no hay ReadAsStringAsync(ct)
                return (await resp.Content.ReadAsStringAsync()).Trim();
            }
        }

        private async Task LaunchUpdaterAsync(CancellationToken ct)
        {
            SetStatus("Descargando actualizador…");
            string updaterPath = Path.Combine(Path.GetTempPath(), "gormatica-updater.exe");
            const string updaterUrl = "https://www.gormatica.com/servicios/soporte-remoto/gormatica-updater.exe";

            try
            {
                await DownloadFileAsync(updaterUrl, updaterPath, delegate (int p)
                {
                    SafeUI(delegate
                    {
                        if (!progreso.Visible) progreso.Visible = true;
                        progreso.Style = ProgressBarStyle.Blocks;
                        progreso.Value = Math.Max(0, Math.Min(p, 100));
                    });
                }, ct);

                if (File.Exists(updaterPath))
                {
                    SetStatus("Lanzando actualizador…");
                    var psi = new ProcessStartInfo
                    {
                        FileName = updaterPath,
                        UseShellExecute = true,
                        Verb = "runas" // Elevar si es necesario
                    };
                    Process.Start(psi);
                    BeginInvoke(new Action(delegate { Application.Exit(); }));
                }
                else
                {
                    SetStatus("No se encontró el actualizador descargado.");
                }
            }
            catch (OperationCanceledException) { }
            catch (UnauthorizedAccessException)
            {
                SetStatus("ERROR: sin permisos para escribir en TEMP.");
            }
            catch (Exception ex)
            {
                SetStatus("ERROR descargando actualizador: " + ex.Message);
            }
        }

        private static async Task DownloadFileAsync(string url, string destination, Action<int> progress, CancellationToken ct)
        {
            using (var resp = await http.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, ct))
            {
                resp.EnsureSuccessStatusCode();

                long? total = resp.Content.Headers.ContentLength;

                using (var input = await resp.Content.ReadAsStreamAsync())
                using (var output = File.Create(destination))
                {
                    byte[] buffer = new byte[81920];
                    long readTotal = 0;
                    int read;
                    while ((read = await input.ReadAsync(buffer, 0, buffer.Length, ct)) > 0)
                    {
                        await output.WriteAsync(buffer, 0, read, ct);
                        readTotal += read;

                        if (total.HasValue && total.Value > 0 && progress != null)
                        {
                            int pct = (int)(readTotal * 100 / total.Value);
                            progress(pct);
                        }
                    }
                    if (progress != null) progress(100);
                }
            }
        }

        // ================== TEAMVIEWER ==================

        private async Task LaunchTeamViewerAsync()
        {
            if (!File.Exists(TeamViewerPath))
            {
                SetStatus("ERROR: No existe " + TeamViewerPath);
                return;
            }

            SafeUI(delegate
            {
                button_teamviewer.Visible = false;
                label_abriendosoporte.Visible = true;
                progreso.Visible = true;
                progreso.Style = ProgressBarStyle.Marquee;
            });

            // Cerrar instancias previas
            SetStatus("Cerrando TeamViewers abiertos…");
            await CloseProcessesByNameAsync(new string[] { "TeamViewer", "TeamViewerQS", "TeamViewer_Service", "TeamViewer_Host", "teamviewer", "teamviewer_service" }, TimeSpan.FromSeconds(5));
            SetStatus(string.Empty);

            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = TeamViewerPath,
                    UseShellExecute = true
                };
                Process.Start(psi);

                bool ok = await WaitUntilAsync(
                    delegate
                    {
                        return Process.GetProcessesByName("TeamViewer").Length > 0
                               || Process.GetProcessesByName("TeamViewerQS").Length > 0;
                    },
                    TimeSpan.FromSeconds(20));

                SafeUI(async delegate
                {
                    if (ok)
                    {
                        progreso.Style = ProgressBarStyle.Blocks;
                        progreso.Value = 100;
                        button_teamviewer.Visible = true;
                    }
                    else
                    {
                        SetStatus("No se pudo confirmar la apertura de TeamViewer.");
                        progreso.Style = ProgressBarStyle.Blocks;
                        progreso.Value = 0;
                        progreso.Visible = false;
                    }
                });
            }
            catch (Exception ex)
            {
                SetStatus("ERROR al abrir TeamViewer: " + ex.Message);
                SafeUI(delegate
                {
                    progreso.Style = ProgressBarStyle.Blocks;
                    progreso.Value = 0;
                    progreso.Visible = false;
                });
            }
        }

        private static async Task CloseProcessesByNameAsync(string[] names, TimeSpan wait)
        {
            foreach (string name in names)
            {
                foreach (var p in Process.GetProcessesByName(name))
                {
                    try
                    {
                        if (!p.HasExited)
                        {
                            p.CloseMainWindow();
                            if (!p.WaitForExit((int)wait.TotalMilliseconds))
                            {
                                // En .NET Framework no existe entireProcessTree
                                p.Kill();
                                p.WaitForExit();
                            }
                        }
                    }
                    catch { }
                }
            }
            await Task.FromResult(0);
        }

        private static async Task<bool> WaitUntilAsync(Func<bool> condition, TimeSpan timeout, int pollMs = 150)
        {
            DateTime start = DateTime.UtcNow;
            while (DateTime.UtcNow - start < timeout)
            {
                if (condition()) return true;
                await Task.Delay(pollMs);
            }
            return false;
        }

        // ================== UTILIDADES UI ==================

        private void SetStatus(string text)
        {
            SafeUI(delegate { if (estado != null) estado.Text = text ?? string.Empty; });
        }

        private void SafeUI(Action action)
        {
            try
            {
                if (IsHandleCreated && InvokeRequired)
                    BeginInvoke(action);
                else
                    action();
            }
            catch { }
        }

        private static void OpenInBrowser(string url)
        {
            try
            {
                var psi = new ProcessStartInfo { FileName = url, UseShellExecute = true };
                Process.Start(psi);
            }
            catch { }
        }



        // ================== EVENTOS ==================
        private void gormatica_Load(object sender, EventArgs e)
        {
            // Puedes dejarlo vacío o mover aquí lógica de arranque si quieres.
        }
        private async void gormatica_Shown(object sender, EventArgs e)
        {

            SafeUI(delegate { if (version != null) version.Text = "Versión " + Application.ProductVersion; });

            // Comprobar actualizaciones sin bloquear UI
            var _ = CheckUpdateAsync();
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            await LaunchTeamViewerAsync();
        }

        private void MenuTecnico_Click(object sender, EventArgs e)
        {
            if (webBrowser1 == null) return;
            webBrowser1.BringToFront();
            webBrowser1.Visible = true;
        }

        private void MenuInicio_Click(object sender, EventArgs e)
        {
            if (webBrowser1 == null) return;
            webBrowser1.Visible = false;
        }


        private void button_anydesk_Click(object sender, EventArgs e)
        {
            OpenInBrowser("https://www.gormatica.com/servicios/soporte-remoto/AnyDesk.exe");
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            if (!_cts.IsCancellationRequested) _cts.Cancel();
        }

        private void button_areacliente_Click(object sender, EventArgs e)
        {
            OpenInBrowser("https://www.gormatica.com/clientes/");
        }
    }
}
