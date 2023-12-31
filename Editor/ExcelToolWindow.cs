using UnityEngine;
using UnityEditor;
using System.Data;
using System.Text;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Lab5Games.ExcelTool.Editor
{
    public class ExcelToolWindow : EditorWindow
    {
        string _excelFilePath = "";

        [MenuItem("Lab5Games/Excel Tool Window")]
        public static void CreateWindow()
        {
            var wnd = GetWindow<ExcelToolWindow>();
            wnd.titleContent = new GUIContent("Excel Tool");
            wnd.maxSize = new Vector2(350, 400);
            wnd.minSize = new Vector2(350, 400);
        }

        private void OnGUI()
        {
            EditorGUILayout.BeginVertical();

            // top
            EditorGUILayout.Space(5);
            EditorGUILayout.BeginHorizontal();
            EditorGUILayout.Space(1);
            
            EditorGUILayout.TextField(_excelFilePath);
            if(GUILayout.Button("Open File")) OpenExcelFile();
            
            EditorGUILayout.Space(1);
            EditorGUILayout.EndHorizontal();
            EditorGUILayout.Space(5);

            // bottom
            EditorGUILayout.Space(25);
            
            if (GUILayout.Button("Export Csv")) ExportCsv();
            if (GUILayout.Button("Export Xml")) ExportXml();
            if (GUILayout.Button("Export Json")) ExportJson();

            EditorGUILayout.EndVertical();
        }

        void OpenExcelFile()
        {
            _excelFilePath = EditorUtility.OpenFilePanel("Selection", Application.dataPath, "xlsx");

            if(!File.Exists(_excelFilePath))
            {
                _excelFilePath = "";
            }

        }

        void ExportCsv()
        {
            if (string.IsNullOrEmpty(_excelFilePath))
                return;

            using(FileStream fileStream = new FileStream(_excelFilePath, FileMode.Open, FileAccess.Read))
            {
                using(ExcelReader excelReader = new ExcelReader(fileStream))
                {
                    StringBuilder strBuilder = new StringBuilder();

                    DataRowCollection rowCollection = excelReader.Read(0);

                    for(int i=0; i<rowCollection.Count; i++)
                    {
                        DataRow row = rowCollection[i];

                        strBuilder.AppendJoin(',', row.ItemArray);

                        if (i < rowCollection.Count - 1) 
                            strBuilder.AppendLine();
                    }


                    string dirName = Path.GetDirectoryName(_excelFilePath);
                    string fileName = Path.GetFileNameWithoutExtension(_excelFilePath);

                    File.WriteAllText(Path.Combine(dirName, fileName) + ".csv", strBuilder.ToString());

                    AssetDatabase.Refresh();

                }
            }
        }

        void ExportXml()
        {
            if (string.IsNullOrEmpty(_excelFilePath))
                return;
        }

        void ExportJson()
        {
            if (string.IsNullOrEmpty(_excelFilePath))
                return;

            using (FileStream fileStream = new FileStream(_excelFilePath, FileMode.Open, FileAccess.Read))
            {
                using (ExcelReader excelReader = new ExcelReader(fileStream))
                {
                    DataRowCollection rowCollection = excelReader.Read(0);

                    JObject jObj = new JObject();
                    JArray rowList = new JArray();

                    for(int i=0; i<rowCollection.Count; i++)
                    {
                        JArray rowData = new JArray();
                        DataRow row = rowCollection[i];

                        for(int j=0; j<row.ItemArray.Length; j++)
                        {
                            rowData.Add(row.ItemArray[j].ToString());
                        }

                        rowList.Add(rowData);
                    }

                    jObj["Rows"] = rowList;

                    string dirName = Path.GetDirectoryName(_excelFilePath);
                    string fileName = Path.GetFileNameWithoutExtension(_excelFilePath);

                    File.WriteAllText(Path.Combine(dirName, fileName) + ".json", jObj.ToString());

                    AssetDatabase.Refresh();
                }
            }
        }
    }
}
