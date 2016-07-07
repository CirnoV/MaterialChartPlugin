using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.ObjectModel;
using System.IO;
using System.Runtime.Serialization;
using System.Xml;
using Livet;
using ProtoBuf;
using Grabacr07.KanColleWrapper;
using MaterialChartPlugin.Properties;

namespace MaterialChartPlugin.Models
{
    public class MaterialLog : NotificationObject
    {
        static readonly string localDirectoryPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "terry_u16", "MaterialChartPlugin");

        public static readonly string ExportDirectoryPath = "MaterialChartPlugin";

        static readonly string saveFileName = "materiallog.dat";

        static string SaveFilePath => Path.Combine(localDirectoryPath, saveFileName);

        private MaterialChartPlugin plugin;

        #region HasLoaded変更通知プロパティ
        private bool _HasLoaded = false;

        public bool HasLoaded
        {
            get
            { return _HasLoaded; }
            set
            { 
                if (_HasLoaded == value)
                    return;
                _HasLoaded = value;
                RaisePropertyChanged();
            }
        }
        #endregion

        public ObservableCollection<TimeMaterialsPair> History { get; private set; }

        public MaterialLog(MaterialChartPlugin plugin)
        {
            this.plugin = plugin;
        }

        public async Task LoadAsync()
        {
            await LoadAsync(SaveFilePath, null);
        }

        private async Task LoadAsync(string filePath, Action onSuccess)
        {
            this.HasLoaded = false;

            if (File.Exists(filePath))
            {
                try
                {
                    using (var stream = File.OpenRead(filePath))
                    {
                        this.History = await Task.Run(() => Serializer.Deserialize<ObservableCollection<TimeMaterialsPair>>(stream));
                    }
                    onSuccess?.Invoke();
                }
                catch (ProtoException ex)
                {
                    plugin.InvokeNotifyRequested(new Grabacr07.KanColleViewer.Composition.NotifyEventArgs(
                        "MaterialChartPlugin.LoadFailed", "로드 실패",
						"자원 데이터 읽기 실패. 데이터가 손상될 수 있습니다."));
                    System.Diagnostics.Debug.WriteLine(ex);
                    if (this.History == null)
                        this.History = new ObservableCollection<TimeMaterialsPair>();
                }
                catch (IOException ex)
                {
                    plugin.InvokeNotifyRequested(new Grabacr07.KanColleViewer.Composition.NotifyEventArgs(
                        "MaterialChartPlugin.LoadFailed", "로드 실패",
						"자원 데이터 읽기 실패. 접근 권한이 있는지 확인해주세요."));
                    System.Diagnostics.Debug.WriteLine(ex);
                    if (this.History == null)
                        this.History = new ObservableCollection<TimeMaterialsPair>();
                }
            }
            else
            {
                if (this.History == null)
                    this.History = new ObservableCollection<TimeMaterialsPair>();
            }

            this.HasLoaded = true;
        }

        public async Task SaveAsync()
        {
            try
            {
                await SaveAsync(localDirectoryPath, SaveFilePath, null);
            }
            catch (IOException ex)
            {
                plugin.InvokeNotifyRequested(new Grabacr07.KanColleViewer.Composition.NotifyEventArgs(
                    "MaterialChartPlugin.SaveFailed", "저장 실패",
					"자원 데이터 저장 실패. 접근 권한이 있는지 확인해주세요."));
                System.Diagnostics.Debug.WriteLine(ex);
            }
        }

        private async Task SaveAsync(string directoryPath, string filePath, Action onSuccess)
        {
            try
            {
                if (!Directory.Exists(directoryPath))
                {
                    Directory.CreateDirectory(directoryPath);
                }

                // オレオレ形式でバイナリ保存とかも考えたけど
                // 今後ネジみたいに新しい資材が入ってくると対応が面倒なのでやめた
                using (var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    await Task.Run(() => Serializer.Serialize(stream, History));
                }

                onSuccess?.Invoke();
            }
            catch (IOException ex)
            {
                plugin.InvokeNotifyRequested(new Grabacr07.KanColleViewer.Composition.NotifyEventArgs(
                    "MaterialChartPlugin.SaveFailed", "저장 실패",
					"자원 데이터 저장 실패. 접근 권한이 있는지 확인해주세요."));
                System.Diagnostics.Debug.WriteLine(ex);
            }
        }

        public async Task ExportAsCsvAsync()
        {
            var csvFileName = CreateCsvFileName(DateTime.Now);
            var csvFilePath = Path.Combine(ExportDirectoryPath, csvFileName);

            try
            {
                if (!Directory.Exists(ExportDirectoryPath))
                {
                    Directory.CreateDirectory(ExportDirectoryPath);
                }

                using (var writer = new StreamWriter(csvFilePath, false, Encoding.UTF8))
                {
                    await writer.WriteLineAsync("시각,연료,탄약,강재,보크사이트,고속수복재,개발자재,고속건조재,개수자재");
                    

                    if (MaterialChartSettings.Default.CsvRecordSummary)
                    {
                        // 日別にまとめる場合
                        var dayList = History
                            .Select(x => new { DateTime = x.DateTime, DateTimeString = x.DateTime.ToShortDateString() })
                            .GroupBy(a => a.DateTimeString)
                            .Select(group => group.Max(a => a.DateTime));

                        if (MaterialChartSettings.Default.CsvRecordOrder)
                        {
                            // 降順にする
                            dayList = dayList.OrderByDescending(x => x);
                        }                           

                        foreach (var day in dayList)
                        {
                            var pair = History.Where(x => x.DateTime.Equals(day)).First();
                            await writer.WriteLineAsync($"{pair.DateTime.ToShortDateString()},{pair.Fuel},{pair.Ammunition},{pair.Steel},{pair.Bauxite},{pair.RepairTool},{pair.DevelopmentTool},{pair.InstantBuildTool},{pair.ImprovementTool}");
                        }
                    }
                    else
                    {
                        // 時刻別
                        var history = History.OrderBy(x => x.DateTime);
                        if (MaterialChartSettings.Default.CsvRecordOrder)
                        {
                            // 降順にする
                            history = history.OrderByDescending(x => x.DateTime);
                        }

                        foreach (var pair in history)
                        {
                            await writer.WriteLineAsync($"{pair.DateTime},{pair.Fuel},{pair.Ammunition},{pair.Steel},{pair.Bauxite},{pair.RepairTool},{pair.DevelopmentTool},{pair.InstantBuildTool},{pair.ImprovementTool}");
                        }
                    }
                }

                plugin.InvokeNotifyRequested(new Grabacr07.KanColleViewer.Composition.NotifyEventArgs(
                    "MaterialChartPlugin.CsvExportCompleted", "내보내기 완료",
                    $"자원 데이터를 저장했습니다: {csvFilePath}")
                {
                    Activated = () =>
                    {
                        System.Diagnostics.Process.Start("EXPLORER.EXE", $"/select,\"\"{csvFilePath}\"\"");
                    }
                });
            }
            catch (IOException ex)
            {
                plugin.InvokeNotifyRequested(new Grabacr07.KanColleViewer.Composition.NotifyEventArgs(
                    "MaterialChartPlugin.ExportFailed", "내보내기 실패",
                    "자원 데이터 저장 실패. 접근 권한이 있는지 확인해주세요."));
                System.Diagnostics.Debug.WriteLine(ex);
            }
        }

        public async Task ImportAsync(string filePath)
        {
            await LoadAsync(filePath, async () =>
                {
                    plugin.InvokeNotifyRequested(new Grabacr07.KanColleViewer.Composition.NotifyEventArgs(
                        "MaterialChartPlugin.ImportComplete", "불러오기 완료",
                        "자원 데이터를 성공적으로 불러왔습니다."));

                    var materials = KanColleClient.Current.Homeport.Materials;
                    History.Add(new TimeMaterialsPair(DateTime.Now, materials.Fuel, materials.Ammunition, materials.Steel,
                        materials.Bauxite, materials.InstantRepairMaterials, materials.DevelopmentMaterials,
                        materials.InstantBuildMaterials, materials.ImprovementMaterials));

                    await SaveAsync();
                }
            );

        }

        public async Task ExportAsync()
        {
            var fileName = CreateExportedFileName(DateTime.Now);
            var filePath = Path.Combine(ExportDirectoryPath, fileName);
            await SaveAsync(ExportDirectoryPath, filePath, () =>
                plugin.InvokeNotifyRequested(new Grabacr07.KanColleViewer.Composition.NotifyEventArgs(
                    "MaterialChartPlugin.ExportComplete", "내보내기 완료",
                    $"자원 데이터를 저장했습니다: {filePath}")
                {
                    Activated = () =>
                    {
                        System.Diagnostics.Process.Start("EXPLORER.EXE", $"/select,\"\"{filePath}\"\"");
                    }
                })
            );
        }

        private string CreateCsvFileName(DateTime dateTime)
        {
            return $"MaterialChartPlugin-{dateTime.ToString("yyMMdd-HHmmssff")}.csv";
        }

        private string CreateExportedFileName(DateTime dateTime)
        {
            return $"MaterialChartPlugin-BackUp-{dateTime.ToString("yyMMdd-HHmmssff")}.dat";
        }
    }
}
