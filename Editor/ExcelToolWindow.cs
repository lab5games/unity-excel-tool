using UnityEngine;
using UnityEditor;
using System.Data;
using System.Text;
using System.IO;
using System.Linq;
using System;

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
                StringBuilder strBuilder = new StringBuilder();

                DataRowCollection rowCollection = DataReader.ReadExcel(fileStream);

                for (int i = 0; i < rowCollection.Count; i++)
                {
                    DataRow row = rowCollection[i];

                    Debug.Log($"{i}: " + string.Join(',', row.ItemArray));

                    strBuilder.AppendJoin(',', row.ItemArray);
                    strBuilder.AppendLine();
                }

                string dirName = Path.GetDirectoryName(_excelFilePath);
                string fileName = Path.GetFileNameWithoutExtension(_excelFilePath);

                File.WriteAllText(Path.Combine(dirName, fileName) + ".csv", strBuilder.ToString());

                AssetDatabase.Refresh();

                Debug.LogWarning("Csv created!");
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
        }
    }
}
