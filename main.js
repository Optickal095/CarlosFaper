const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");
const cvs = require("csv-parser");

// Ruta .xlsx
const filePath = path.join(__dirname, "./2020-05-01-CasosConfirmados.xlsx");

// Exports (Informe 2 - 4)
const outputFilePathComuna = path.join(__dirname, "./Casos_Comuna.csv");
const outputFilePathRegion = path.join(__dirname, "./Casos_Region.csv");
const cvsData = [];
const cvsDataRegion = [];

const workbook = XLSX.readFile(filePath);

const sheetNames = workbook.SheetNames;

let totalCasosConfirmados = 0;
let totalFilas = 0;
let casosPorComuna = {};
let casosPorRegion = {};

// Se leen los datos
sheetNames.forEach((sheetName) => {
  const worksheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

  // Se rescatan "Casos Confirmados de Índice de tabla"
  const newData = data.map((row) => row[0].split(","));
  const casosConfirmadosIndex = newData[0].indexOf("Casos Confirmados");

  // Iteración para Informes
  newData.slice(1).forEach((row) => {
    const comuna = row[2];
    const region = row[0];

    //Total de casos confirmados
    const casosConfirmados = parseFloat(row[casosConfirmadosIndex]);

    //Inicialización Casos Por Comuna
    if (!isNaN(casosConfirmados)) {
      if (!casosPorComuna[comuna]) {
        casosPorComuna[comuna] = 0;
      }

      //Inicialización Casos Por Region
      if (!casosPorRegion[region]) {
        casosPorRegion[region] = 0;
      }

      // Cálculos de casos confirmados por region/comuna/total/promedio
      casosPorRegion[region] += casosConfirmados;
      casosPorComuna[comuna] += casosConfirmados;
      totalCasosConfirmados += casosConfirmados;
      totalFilas++;
    }
  });

  // Calcular promedio de contagios totales
  calcularPromedioCasosConfirmados = totalCasosConfirmados / totalFilas;

  // Comuna con más contagios
  const comunaConMasContagios = Object.keys(casosPorComuna).reduce(
    (comunaMax, comuna) => {
      return casosPorComuna[comuna] > casosPorComuna[comunaMax]
        ? comuna
        : comunaMax;
    },
    Object.keys(casosPorComuna)[0]
  );

  // Region con más contagios
  const regionConMasContagios = Object.keys(casosPorRegion).reduce(
    (regionMax, region) => {
      return casosPorRegion[region] > casosPorRegion[regionMax]
        ? region
        : regionMax;
    },
    Object.keys(casosPorRegion)[0]
  );

  // Contagios máximos por comuna/region
  const contagiosMaximosComuna = casosPorComuna[comunaConMasContagios];
  const contagiosMaximosRegion = casosPorRegion[regionConMasContagios];

  // Total de casos confirmados (Informe 1)
  console.log(`Total de casos confirmados: ${totalCasosConfirmados}`);

  // Promedio de total de casos confirmados (Informe 1)
  console.log(`Promedio: ${calcularPromedioCasosConfirmados}`);

  // Casos totales de contagios por Comuna (Informe 2)
  console.log("Casos por comuna: ");
  for (const comuna in casosPorComuna) {
    console.log(`${comuna}: ${casosPorComuna[comuna]}`);
    cvsData.push({
      Comuna: comuna,
      CasosConfirmados: casosPorComuna[comuna],
    });
  }

  // Creación de Casos_Comuna.csv (Informe 2)
  fs.writeFileSync(outputFilePathComuna, "Comuna,CasosConfirmados\n");

  cvsData.forEach((row) => {
    fs.appendFileSync(
      outputFilePathComuna,
      `${row.Comuna},${row.CasosConfirmados}\n`
    );
  });

  // Comuna con más contagios (Informe 3)
  console.log(
    `La comuna con más contagiados: ${comunaConMasContagios}. Número de contagios: ${contagiosMaximosComuna}`
  );

  // Region con más contagios (Informe 4)
  for (const region in casosPorRegion) {
    cvsDataRegion.push({
      Region: region,
      CasosConfirmados: casosPorRegion[region],
    });
  }

  // Región con más contagios (Informe 4)
  console.log(
    `La region con más contagiados: ${regionConMasContagios}. Número de contagios: ${contagiosMaximosRegion}`
  );

  // Creación de Casos_Region.csv
  fs.writeFileSync(outputFilePathRegion, "Region,CasosConfirmados\n");

  cvsDataRegion.forEach((row) => {
    fs.appendFileSync(
      outputFilePathRegion,
      `${row.Region},${row.CasosConfirmados}\n`
    );
  });
});
