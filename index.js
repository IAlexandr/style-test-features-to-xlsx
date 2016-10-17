import { FeatureServer } from 'arcgis-api-client';
import {featuresToWorkBook} from './lib/utils';
import XLSX from 'xlsx-style';

const featureServerUrl = 'https://***.ru/arcgis/rest/services/pomesheniya/nezhil_pom_v8/FeatureServer/0';
const username = '****';
const password = '****';
const sheetName = 'Отчет';

const featureServer = new FeatureServer({ featureServerUrl, username, password });
featureServer.query({ featureServerUrl, username, password }, { returnGeometry: false }).then(featuresInfo => {
  const wb = featuresToWorkBook({featuresInfo, sheetName});
  XLSX.writeFile(wb, 'report.xlsx', {
    bookType: 'xlsx',
    bookSST: false,
    type: 'binary',
    cellStyles: true
  });
  process.exit(0);
}).catch(err => {
  console.log(err.message);
});
