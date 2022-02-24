// Sheet name where data is stored
var SHEET_NAME = "ruedas";

var SCRIPT_PROP = PropertiesService.getScriptProperties(); // new property service

// This function glues spreadsheet and Apps Script project
function setup() {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  SCRIPT_PROP.setProperty("key", doc.getId());
}

function doGet() {
  var output = HtmlService.createTemplateFromFile('Index').evaluate();
  output.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return output
}

// Creates an import or include function so files can be added inside the main index.
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
};

// Find total tyre ratings
function findTotalRatings(tyreName = "") {
  var sheet = SpreadsheetApp.getActiveSheet();
  var column = 2;
  // GET tyre model ocurrences
  var dataFromSheetName = [];
  dataFromSheetName = sheet.getRange(2, column, sheet.getLastRow(), 1).getValues();
  dataFromSheetName = dataFromSheetName.filter(item => item);

  //Get data from column and parse it to Chart.js format
  var ocurrences = 0;
  for (row in dataFromSheetName) {
    if (dataFromSheetName[row][0] !== "" && tyreName === dataFromSheetName[row][0]) {
      ocurrences = ocurrences + 1;
    }
  }
  return ocurrences;
}

// Find row by given value
function findByNameType(tyreName = "", searchedValue = "") {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var columnCount = sheet.getDataRange().getLastColumn();

  var i = data.flat().indexOf(tyreName);
  var columnIndex = i % columnCount
  var rowIndex = ((i - columnIndex) / columnCount);

  var value = [];
  if (searchedValue === "tyreType") {
    var tyreType = SpreadsheetApp.getActiveSheet().getRange(rowIndex + 1, 18).getValue();
    value.push(tyreType);

    if (tyreType === "Circuito") {
      value.push("Los neumáticos de circuito son ruedas diseñadas únicamente para la competición y los circuitos. Proporcionan un agarre máximo pero, su temperatura de funcionamiento es alta. La durabilidad en kilometraje es muy baja y se recomienda su uso solo para usuarios habituales de los circuitos.");
    } else if (tyreType === "Sport") {
      value.push("Estos neumáticos están diseñados para motos de calle y motos deportivas de alto rendimiento, tienen un peso más ligero y un excelente agarre. Su durabilidad en kilometraje suele ser baja y su temperatura de funcionamiento óptimo suele ser algo más elevada que un neumático estándar. Se recomienda que las utilicen usuarios cuyo uso principal sean rutas y/o circuitos.");
    }
    else if (tyreType === "Touring") {
      value.push("Los neumáticos Cruiser o Touring ofrecen un gran kilometraje y una buena tracción en condiciones de humedad. Este tipo de neumáticos busca un equilibrio en su comportamiento, para ser usado eficazmente bajo cualquier circunstancia. Su uso principal es para usuarios que usen su moto día a día para cualquier trayecto o que suelan realizar viajes largos con sus motos.");
    }
    else if (tyreType === "Adventure") {
      value.push("Los neumáticos de Adventure o Trail son usados por aquellos usuarios que quieran experimentar el off-roading junto con la posibilidad de circular por carreteras asfaltadas. Se recomienda su uso a aquellas personas que ocasionalmente salgan de las carreteras convencionales, pero que también hagan uso de ellas. Se les podría llamar ruedas todoterreno.");
    }
    else if (tyreType === "Scooter") {
      value.push("Estos neumáticos están diseñados para motos Scooter, dentro de los neumáticos para este tipo de motos podemos encontrar distintos compuestos, durabilidad o agarre. Su característica principal es el diseño de los mismos, ya que están hechos para ruedas con un diámetro más pequeño. Su agarre u durabilidad dependerán del modelo elegido, por ello es importante consular las gráficas que aparecen a continuación.");
    }
    else if (tyreType === "Custom") {
      value.push("Estos neumáticos están diseñados para motos Custom, tienen un buen compromiso entre durabilidad y agarre. Son estables y por norma general aguantan bien el peso de las Custom más pesadas. Su principal uso serian carretera, rutas o ciudad pero sin pretenciones deportivas.");
    } else {
      value.push("No disponemos de información acerca de este tipo de neumático.");
    }
  } else {
    value = "Value not found"
  }
  return i >= 0 ? value : "Value not found";
}

// Get values from sheet
function getDataFromSheet(wanted_data = "", model_selected = null) {
  try {
    let doc = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
    let sheet = doc.getSheetByName(SHEET_NAME);
    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Find column wanted
    if (wanted_data === "tyreModels") {
      // GET tyre model names
      var dataFromSheetName = [];
      dataFromSheetName = sheet.getRange(2, 2, sheet.getLastRow(), 1).getValues();
      dataFromSheetName = dataFromSheetName.filter(item => item);

      //Get data from column and parse it to Chart.js format
      var modelNames = [];
      for (row in dataFromSheetName) {
        if (dataFromSheetName[row][0] !== "") {
          var modelName = dataFromSheetName[row][0];
          // Check if model is already in array
          if (!modelNames.includes(modelName)) {
            modelNames.push(modelName);
          }
        }
      }

      // Alphabetical order
      modelNames.sort();
      return modelNames;

    } else if (wanted_data === "tyreUse") {
      // GET tyre model uses
      var dataFromSheetName = [];
      var dataFromSheetCircuito = [];
      var dataFromSheetCiudad = [];
      var dataFromSheetRutas = [];
      var dataFromSheetOffroad = [];

      dataFromSheetName = sheet.getRange(2, 2, sheet.getLastRow(), 1).getValues();
      dataFromSheetCircuito = sheet.getRange(2, 3, sheet.getLastRow(), 1).getValues();
      dataFromSheetCiudad = sheet.getRange(2, 4, sheet.getLastRow(), 1).getValues();
      dataFromSheetRutas = sheet.getRange(2, 5, sheet.getLastRow(), 1).getValues();
      dataFromSheetOffroad = sheet.getRange(2, 6, sheet.getLastRow(), 1).getValues();
      dataFromSheetAutovia = sheet.getRange(2, 16, sheet.getLastRow(), 1).getValues();

      //Get data from each column
      var tyreModelsUseTotal = 0;
      var useCircuito = 0;
      var useCiudad = 0;
      var useRutas = 0;
      var useOffroad = 0;
      var useAutovia = 0;
      for (row in dataFromSheetName) {
        if (dataFromSheetName[row][0] !== "" && model_selected === dataFromSheetName[row][0]) {
          useCircuito = useCircuito + dataFromSheetCircuito[row][0];
          useCiudad = useCiudad + dataFromSheetCiudad[row][0];
          useRutas = useRutas + dataFromSheetRutas[row][0];
          useOffroad = useOffroad + dataFromSheetOffroad[row][0];
          useAutovia = useAutovia + dataFromSheetAutovia[row][0];

          tyreModelsUseTotal = tyreModelsUseTotal + 1;
        }
      }
      // Average
      useCircuito = useCircuito / tyreModelsUseTotal;
      useCiudad = useCiudad / tyreModelsUseTotal;
      useRutas = useRutas / tyreModelsUseTotal;
      useOffroad = useOffroad / tyreModelsUseTotal;
      useAutovia = useAutovia / tyreModelsUseTotal;

      return [useCircuito, useRutas, useCiudad, useAutovia, useOffroad];

    } else if (wanted_data === "tyreGrip") {
      // GET tyre model grip
      var dataFromSheetName = [];
      var dataFromSheetSeco = [];
      var dataFromSheetMojado = [];
      var dataFromSheetTierra = [];
      var dataFromSheetNieve = [];

      dataFromSheetName = sheet.getRange(2, 2, sheet.getLastRow(), 1).getValues();
      dataFromSheetSeco = sheet.getRange(2, 8, sheet.getLastRow(), 1).getValues();
      dataFromSheetMojado = sheet.getRange(2, 9, sheet.getLastRow(), 1).getValues();
      dataFromSheetTierra = sheet.getRange(2, 10, sheet.getLastRow(), 1).getValues();
      dataFromSheetNieve = sheet.getRange(2, 11, sheet.getLastRow(), 1).getValues();

      //Get data from each column
      var tyreModelsGripTotal = 0;
      var useSeco = 0;
      var useMojado = 0;
      var useTierra = 0;
      var useNieve = 0;
      for (row in dataFromSheetName) {
        if (dataFromSheetName[row][0] !== "" && model_selected === dataFromSheetName[row][0]) {
          useSeco = useSeco + dataFromSheetSeco[row][0];
          useMojado = useMojado + dataFromSheetMojado[row][0];
          useTierra = useTierra + dataFromSheetTierra[row][0];
          useNieve = useNieve + dataFromSheetNieve[row][0];

          tyreModelsGripTotal = tyreModelsGripTotal + 1;
        }
      }
      useSeco = useSeco / tyreModelsGripTotal;
      useMojado = useMojado / tyreModelsGripTotal;
      useTierra = useTierra / tyreModelsGripTotal;
      useNieve = useNieve / tyreModelsGripTotal;
      
      return [useSeco, useMojado, useTierra, useNieve];
    } else if (wanted_data === "tyreBreak") {
      // GET tyre model break rating
      var dataFromSheetName = [];
      var dataFromSheetFrenadaSeco = [];
      var dataFromSheetFrenadaMojado = [];

      dataFromSheetName = sheet.getRange(2, 2, sheet.getLastRow(), 1).getValues();
      dataFromSheetFrenadaSeco = sheet.getRange(2, 14, sheet.getLastRow(), 1).getValues();
      dataFromSheetFrenadaMojado = sheet.getRange(2, 15, sheet.getLastRow(), 1).getValues();

      //Get data from each column
      var tyreModelsBreakTotal = 0;
      var breakSeco = 0;
      var breakMojado = 0;
      for (row in dataFromSheetName) {
        if (dataFromSheetName[row][0] !== "" && model_selected === dataFromSheetName[row][0]) {
          breakSeco = breakSeco + dataFromSheetFrenadaSeco[row][0];
          breakMojado = breakMojado + dataFromSheetFrenadaMojado[row][0];

          tyreModelsBreakTotal = tyreModelsBreakTotal + 1;
        }
      }

      // Average
      breakSeco = breakSeco / tyreModelsBreakTotal
      breakMojado = breakMojado / tyreModelsBreakTotal

      return [breakSeco, breakMojado];
    } else if (wanted_data === "tyreRating") {
      // GET tyre model rating
      var dataFromSheetName = [];
      var dataFromSheetRating = [];

      dataFromSheetName = sheet.getRange(2, 2, sheet.getLastRow(), 1).getValues();
      dataFromSheetRating = sheet.getRange(2, 12, sheet.getLastRow(), 1).getValues();

      //Get data from each column
      var tyreModelsRatingTotal = 0;
      var tyreRating = 0;
      for (row in dataFromSheetName) {
        if (dataFromSheetName[row][0] !== "" && model_selected === dataFromSheetName[row][0]) {
          tyreRating = tyreRating + dataFromSheetRating[row][0];
          tyreModelsRatingTotal = tyreModelsRatingTotal + 1;
        }
      }

      return tyreRating / tyreModelsRatingTotal;

    } else if (wanted_data === "tyreCostVal") {
      // GET tyre cost value
      var dataFromSheetName = [];
      var dataFromSheetCostval = [];

      dataFromSheetName = sheet.getRange(2, 2, sheet.getLastRow(), 1).getValues();
      dataFromSheetCostval = sheet.getRange(2, 13, sheet.getLastRow(), 1).getValues();

      //Get data from each column
      var tyreModelsCostValTotal = 0;
      var tyreCostVal = 0;
      for (row in dataFromSheetName) {
        if (dataFromSheetName[row][0] !== "" && model_selected === dataFromSheetName[row][0]) {
          tyreCostVal = tyreCostVal + dataFromSheetCostval[row][0];
          tyreModelsCostValTotal = tyreModelsCostValTotal + 1;
        }
      }

      return tyreCostVal / tyreModelsCostValTotal;

    } else if (wanted_data === "tyreDuration") {
      // GET tyre model uses
      var dataFromSheetName = [];
      var dataFromSheetDuration = [];

      dataFromSheetName = sheet.getRange(2, 2, sheet.getLastRow(), 1).getValues();
      dataFromSheetDuration = sheet.getRange(2, 7, sheet.getLastRow(), 1).getValues();

      //Get data from each column
      var tyreModelsDurationTotal = 0;
      var tyreDuration = 0;
      for (row in dataFromSheetName) {
        if (dataFromSheetName[row][0] !== "" && model_selected === dataFromSheetName[row][0]) {
          tyreDuration = tyreDuration + dataFromSheetDuration[row][0];
          tyreModelsDurationTotal = tyreModelsDurationTotal + 1;
        }
      }

      return tyreDuration / tyreModelsDurationTotal;

    }else if (wanted_data === "tyreAvgPrice") {
      // GET tyre model uses
      var dataFromSheetName = [];
      var dataFromSheetAvgPrice = [];

      dataFromSheetName = sheet.getRange(2, 2, sheet.getLastRow(), 1).getValues();
      dataFromSheetAvgPrice = sheet.getRange(2, 17, sheet.getLastRow(), 1).getValues();

      //Get data from each column
      var tyreModelPrices= 0;
      var tyreModelPricesTotal = 0;
      for (row in dataFromSheetName) {
        if (dataFromSheetName[row][0] !== "" && model_selected === dataFromSheetName[row][0]) {
          tyreModelPrices = tyreModelPrices + dataFromSheetAvgPrice[row][0];
          tyreModelPricesTotal = tyreModelPricesTotal + 1;
        }
      }

      return tyreModelPrices / tyreModelPricesTotal;

    }

  }
  catch (e) {
    return []
  }

}