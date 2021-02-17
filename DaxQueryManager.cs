using System;
using System.Collections.Generic;
using System.Text;
using TOM = Microsoft.AnalysisServices.Tabular;
using Microsoft.AnalysisServices.AdomdClient;
using System.IO;
using System.Diagnostics;

namespace DaxQueryExporter {

  class DaxQueryManager {

    private static string GetConnectStringForUser() {
      string workspaceConnection = "powerbi://api.powerbi.com/v1.0/myorg/YOUR_PREMIUM_WORKSPACE";
      string dataset = "YOUR_DATASET";
      string userId = "YOUR_USER_ACCOUNT";
      string password = "YOUR_USER_PASSWORD";
      return $"DataSource={workspaceConnection};Catalog={dataset};User ID={userId};Password={password};";
    }

    private static string GetConnectStringForServicePrincipal() {
      string workspaceConnection = "powerbi://api.powerbi.com/v1.0/myorg/YOUR_PREMIUM_WORKSPACE";
      string dataset = "YOUR_DATASET";
      string tenantId = "YOUR_TENANT_ID";
      string appId = "YOUR_APP_ID";
      string appSecret = "YOUR_APP_SECRET";
      return $"DataSource={workspaceConnection};Catalog={dataset};User ID=app:{appId}@{tenantId};Password={appSecret};";
    }

  public static void ConvertDaxQueryToCsv(string DaxQuery, string FileName) {

      string connectString = GetConnectStringForUser();
      AdomdConnection adomdConnection = new AdomdConnection(connectString);
      adomdConnection.Open();

      AdomdCommand adomdCommand = new AdomdCommand(DaxQuery, adomdConnection);
      AdomdDataReader reader = adomdCommand.ExecuteReader();

      ConvertReaderToCsv(FileName, reader);

      reader.Dispose();
      adomdConnection.Close();

    }

    private static void ConvertReaderToCsv(string FileName, AdomdDataReader Reader, bool OpenInExcel = true) {

      string csv = string.Empty;

      for (int col = 0; col < Reader.FieldCount; col++) {
        csv += Reader.GetName(col);
        csv += (col < (Reader.FieldCount - 1)) ? "," : "\n";
      }

      // Create a loop for every row in the resultset
      while (Reader.Read()) {
        // Create a loop for every column in the current row
        for (int i = 0; i < Reader.FieldCount; i++) {
          csv += Reader.GetValue(i);
          csv += (i < (Reader.FieldCount - 1)) ? "," : "\n";
        }
      }

      string filePath = System.IO.Directory.GetCurrentDirectory() + @"\" + FileName;
      StreamWriter writer = File.CreateText(filePath);
      writer.Write(csv);
      writer.Flush();
      writer.Dispose();

      if (OpenInExcel) {
        OpenCsvInExcel(filePath);
      }

    }

    private static void OpenCsvInExcel(string FilePath) {

      ProcessStartInfo startInfo = new ProcessStartInfo();

      bool excelFound = false;
      if (File.Exists("C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE")) {
        startInfo.FileName = "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE";
        excelFound = true;
      }
      else {
        if (File.Exists("C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\EXCEL.EXE")) {
          startInfo.FileName = "C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\EXCEL.EXE";
          excelFound = true;
        }
      }
      if (excelFound) {
        startInfo.Arguments = FilePath;
        Process.Start(startInfo);
      }
      else {
        System.Console.WriteLine("Coud not find Microsoft Exce on this PC.");
      }

    }

  }


}
