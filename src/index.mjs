import axios from 'axios';
import xl from 'excel4node';
import { capitalize } from './utils/string.mjs';

const requestCountries = async () => {
  const baseUrl = 'https://restcountries.com/v3.1/all'

  const response = await axios.get(baseUrl);

  if (response.status !== 200)
    throw Error('Request Fail')

  let data = []

  if (response.data) {
    // Normalize Data
    data = response.data.map((data) => ({
      name: data?.name?.common,
      capital: data?.capital || '-',
      area: data?.area || '-',
      currencies: data?.currencies ? Object.keys(data?.currencies).join() : '-'
    }))
  }

  return data;
}

const createXMLFile = async (data, header = '') => {
  if (!data)
    throw Error('Data Array is required')

  const workBook = new xl.Workbook();
  var workSheet = workBook.addWorksheet(header);

  //Header Cell
  workSheet.cell(1, 1, 1, 4, true).string(header).style({
    font: {
      size: 16,
      bold: true,
      color: '#4F4F4F',
    },
    alignment: {
      horizontal: ['centerContinuous',]
    }
  })

  //Create Columns
  Object.keys(data[0]).forEach((name, index) => {
    const col = index + 1;
    workSheet.cell(2, col).string(capitalize(`${name}`)).style({
      font: {
        size: 12,
        bold: true,
        color: '#808080',
      },

    })
  })

  //Create Rows
  data.forEach(async (country, rowIndex) => {
    const firstRow = 3;
    try {
      Object.values(country)
        .forEach((value, columIndex) => {
          const cell = workSheet.cell(firstRow + rowIndex, columIndex + 1).style({
            font: {
              size: 12,
            },
            numberFormat: '#,##0.00; (#,##0.00); -',
          })

          if (typeof value === 'number') {
            cell.number(value)
          } else {
            cell.string(value)
          }
        }

        )
    } catch (error) {
      console.error(error)
    }
  })

  workBook.write('./outputs/countries.xlsx');
}


async function main() {
  const contries = await requestCountries()

  createXMLFile(contries, 'Countries List');
}


main()