const express = require("express")
const excel = require("exceljs")

const xlsx = require('xlsx')
const excelRoute = express.Router()

var dataObject = require('./dataObject')
const text = require('./text')

var fs = require("fs")
const { json } = require("body-parser")

//******** get data from excel **********/
async function getDataFromFile(){
  try{
      const file = xlsx.readFile('data.xlsx')
      let wsheets = {}
      for( let sheetname of file.SheetNames ){
        wsheets[sheetname] = xlsx.utils.sheet_to_json(file.Sheets[sheetname])
      }  
      return JSON.stringify(wsheets)

  }
  catch(error){
    return (error)
  }

}

//*****  create  *******/
excelRoute.get('/export', async (request, response) => {
    try{ 

      // data from excel file, gets value to fill the new table
        const dataList = await getDataFromFile().then(value => JSON.parse(value))
        
        const workbook = new excel.Workbook()
        workbook.creator = 'AndreaTopic'
        workbook.created = new Date(2023, 5, 16)

        const sheet = workbook.addWorksheet("nalog")    
        
        sheet.mergeCells('A6:I11')

        sheet.getCell('A5').value = `Predmet: ${dataObject.Predmet}`,{
            font: {
              color: {
                argb: '00FF0000',
                theme: 1,
              },
            }
          }
         
        sheet.getCell('A6').value = text
        sheet.mergeCells('A12:B12')
        sheet.getCell('A12').value = "Katedra"
        sheet.getCell('C12').value = "Studij"
        sheet.getCell('D12').value = "ak. god."
        sheet.getCell('E12').value = "stud. god."
        sheet.getCell('F12').value = "početak turnusa"
        sheet.getCell('G12').value = "kraj turnusa"
        sheet.mergeCells('H12:I12')
        sheet.getCell('H12').value = "br sati predviđen programom"
        sheet.mergeCells('A13:B13')
        sheet.getCell('A13').value = `${dataObject.Katedra} `
        sheet.getCell('C13').value = `${dataObject.Studiji} `
        sheet.getCell('D13').value = `${dataObject.akgod} `
        sheet.getCell('E13').value = `${dataObject.studgod} `
        sheet.getCell('F13').value = `${dataObject.početakturnusa} `
        sheet.getCell('G13').value = `${dataObject.krajturnusa} `
        sheet.getCell('H13').value = `${dataObject.brsati} `

        sheet.mergeCells('E15:G15')
        sheet.mergeCells('K15:M15')
        sheet.mergeCells('A15:A16')
        sheet.mergeCells('A25:C25')
        sheet.mergeCells('A28:C29')
        sheet.mergeCells('A34:C35')
        sheet.mergeCells('J34:L35')
        sheet.mergeCells('B15:B16')
        sheet.mergeCells('C15:C16')
        sheet.mergeCells('D15:D16')
        sheet.mergeCells('I15:I16')
        sheet.mergeCells('J15:J16')
        sheet.mergeCells('H15:H16')
        sheet.mergeCells('N15:N16')
        sheet.getCell('E15').value = "Sati nastave"
        sheet.getCell('K15').value = "Bruto iznos"

        sheet.getCell('E16').value = "pred"
        sheet.getCell('F16').value = "sem"
        sheet.getCell('G16').value = "vjež"
        sheet.getCell('K16').value = "pred"
        sheet.getCell('L16').value = "sem"
        sheet.getCell('M16').value = "vjež"

        sheet.getCell('A15').value = "Redni broj"
        sheet.getCell('B15').value = "Nastavnik/Suradnik"
        sheet.getCell('C15').value = "Zvanje"
        sheet.getCell('D15').value = "Status"
        sheet.getCell('I15').value = "Bruto satnica predavanja (EUR)"
        sheet.getCell('J15').value = "Bruto satnica seminari (EUR)"
        sheet.getCell('H15').value = "Bruto satnica vježbe (EUR)"
        sheet.getCell('N15').value = "Ukupno za isplatu (EUR)"
        sheet.getCell('A25').value = "UKUPNO"


        var len = (dataObject.tablica.length)
        sheet.getCell('D25').value = `${dataObject.tablica.length}`
        for(let i=0; i <= len; i++){
            let start = 17
            sheet.getCell("A"+start.toString()).value = `${start}`
            sheet.getCell("B"+start.toString()).value  = `${dataObject.tablica[i]}`
            //sheet.getCell(`{start}`).value = `${dataObject.tablica[i].Zvanje}`
           // sheet.getCell(`D${start}`).value = `${dataObject.tablica[i].Status}`
           // sheet.getCell(`E${start}`).value = `${dataObject.tablica[i].satin.pred}`
           // sheet.getCell(`F${start}`).value = `${dataObject.tablica[i].satin.sem}`
           // sheet.getCell(`G${start}`).value = `${dataObject.tablica[i].satin.vjez}`
           // sheet.getCell(`K${start}`).value = `${dataObject.tablica[i].brutoiznos.pred}`
           // sheet.getCell(`L${start}`).value = `${dataObject.tablica[i].brutoiznos.sem}`
           // sheet.getCell(`M${start}`).value = `${dataObject.tablica[i].brutoiznos.vjez}`
           // sheet.getCell(`N${start}`).value = `${dataObject.tablica[i].UkupnoisplatE}`

            start++

        }


        sheet.getCell('A34').value = `Prodekan za financije i upravljanje \n Prof. dr. sc. ${dataObject.prodekan1} `
        sheet.getCell('A34').value = `Prodekanica za nastavu i studentska pitanja
            Prof. dr. sc.  ${dataObject.prodekan2} `
        sheet.getCell('J34').value = `Dekan
            Prof. dr. sc. ${dataObject.dekan} `
        
    
        //return response.status(200).send(JSON.stringify(dataObject))
        response.setHeader(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          );
        response.setHeader(
            "Content-Disposition",
            "attachment;filename= "+"nalog.xlsx"
          );  
        
          return workbook.xlsx.write(response)
        
    }
    catch (error) {
        console.log(error)
        return response.status(500).send(error)
    }
    
  });





module.exports = excelRoute