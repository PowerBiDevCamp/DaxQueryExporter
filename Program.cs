using System;

namespace DaxQueryExporter {
  class Program {
    static void Main() {

      DaxQueryManager.ConvertDaxQueryToCsv(Properties.Resources.SalesByState_dax, "SalesByState.csv");

    }
  }
}
