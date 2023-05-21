const express = require("express")
const excel = require("exceljs")

const xlsx = require('xlsx')
const excelRoute = express.Router()

var testData = require('./dataObject')
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
//***** get data for first table and eliminate duplicates *******

function findData(arr){
  var newArr = []
  var valueArr = arr.map(function(item){ return item.PredmetNaziv });
  var isDuplicate = valueArr.some(function(item, idx){ 
    return valueArr.indexOf(item) != idx 
  })
  for(var i=0; i < arr.length ; i++ ){
    newArr.push(arr[i])
    if(isDuplicate){
      return newArr
    }
  }
  console.log(newArr,isDuplicate)
return newArr

}

// ******calculate sum *******

function ukupno(arr){
  let newArr = {"PlaniraniSatiPredavanja":0, "PlaniraniSatiSeminari":0, "PlaniraniSatiVjezbe":0}
  var predavanjaArr = arr.map(function(item){ return item.PlaniraniSatiPredavanja })
  var seminariArr = arr.map(function(item){ return item.PlaniraniSatiSeminari })
  var vjezbeArr = arr.map(function(item){ return item.PlaniraniSatiVjezbe })
  let sum = 0
  predavanjaArr.forEach(item => {
      sum += item
  })
  newArr.PlaniraniSatiPredavanja = sum
  sum = 0

  seminariArr.forEach(item => {
    sum += item
  })
  newArr.PlaniraniSatiSeminari = sum
  sum = 0
  vjezbeArr.forEach(item => {
    sum += item
  })
  newArr.PlaniraniSatiVjezbe = sum
  return newArr
}

//*****  create  route *******/
excelRoute.get('/export', async (request, response) => {
    try{ 

      // data from excel file, gets value to fill the new table
        const dataJSON = await getDataFromFile().then(value => JSON.parse(value))
        const dataList = dataJSON.List1
        const doubleData = findData(dataList)
        const total = ukupno(dataList)
        //console.log(doubleData)
        
        const workbook = new excel.Workbook()
        workbook.creator = 'AndreaTopic'
        workbook.created = new Date(2023, 5, 16)

        const sheet = workbook.addWorksheet("nalog")  
        
                // ****************style *****************//

                function borderCellHeader(cellStr){
                  sheet.getCell(cellStr).fill = {
                    type: 'pattern',
                    pattern:'mediumGray',
                    fgColor:{argb:'D3D3D3'}
                  }
                  
                  sheet.getCell(cellStr).alignment = { wrapText: true,  vertical: "middle", horizontal: "center" }
                
                  sheet.getCell(cellStr).height = 50
                  sheet.getCell(cellStr).style.font = { bold: true }
                  return sheet.getCell(cellStr).border = {
                        top: {style:'thick', color: {argb:'#000000'}},
                        left: {style:'thick', color: {argb:'#000000'}},
                        bottom: {style:'thick', color: {argb:'#000000'}},
                        right: {style:'thick', color: {argb:'#000000'}}
                      }
                    }
   //_________________________________________________________________________________________     
                function borderCell(cellStr){
                  sheet.getCell(cellStr).value = {
                    richText: [
                      {
                        font: {
                          color: {
                            argb: '00FF0000',
                            theme: 1,
                          },
                        },
                      },
                    ],
                  }
                  sheet.getCell(cellStr).alignment = {  wrapText: true,  vertical: "middle" , horizontal: "center" }
                      return sheet.getCell(cellStr).border = {
                        top: {style:'thick', color: {argb:'#000000'}},
                        bottom: {style:'thick', color: {argb:'#000000'}}
                       }
                    }
    //_________________________________________________________________________________________          
       
    
      sheet.getRow('16').height = 50
      sheet.mergeCells('A6:I11')
      sheet.getCell('A5').value = {
          richText: [
            {
              text: 'Predmet:',
            },
            {
              text: ` ${doubleData[0].PredmetNaziv} ${doubleData[0].PredmetKratica}`,
              font: {
                color: {
                  argb: '00FF0000',
                  theme: 1,
                },
              },
            },
          ],
        }

        sheet.getRow('12').height = 50 
        
        sheet.getCell('A6').alignment = { wrapText: true , horizontal: 'center' }
        sheet.mergeCells('H13:I13')
        var headerCells = ['A12','B12','C12','D12','E12','F12','G12','H12','I12','A15','B15','C15','D15','E15','E16','F16','G16','K16','L16','M16','H15','I15','J15','K15','N15']
        headerCells.forEach(element => borderCellHeader(element))
        var subjectCells = ['A13','B13','C13','D13','E13','F13','G13','H13','I13']
        subjectCells.forEach(element => borderCell(element))
        sheet.getCell('H13').border = {right: {style:'thick', color: {argb:'#000000'}}, bottom: {style:'thick', color: {argb:'#000000'}} }
        sheet.getCell('A13').font = { color: {argb:'#FF0000'} }
        
        sheet.getCell('A6').value = sheet.getCell('A13').value = {
          richText: [
            {
              text: 'NALOG ZA ISPLATU ',
              font: {
               bold:true,
               size:14
               },
            },
            {
              text: text,
            },
          ],
        }

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
        sheet.getCell('A13').value = {
          richText: [
            {
              text: `${doubleData[0].Katedra} `,
              font: {
                color: {
                  argb: '00FF0000',
                  theme: 1,
                },
              },
            },
          ],
        }
        sheet.getCell('A13').alignment = { wrapText: true , horizontal: 'center' }
        sheet.getCell('C13').value = {
          richText: [
            {
              text: `${doubleData[0].Studij} `,
              font: {
                color: {
                  argb: '00FF0000',
                  theme: 1,
                },
              },
            },
          ],
        } 
        sheet.getCell('D13').value ={
          richText: [
            {
              text: `${doubleData[0].SkolskaGodinaNaziv} `,
              font: {
                color: {
                  argb: '00FF0000',
                  theme: 1,
                },
              },
            },
          ],
        }
        sheet.getCell('E13').value = {
          richText: [
            {
              text: `${doubleData[0].PkSkolskaGodina} `,
              font: {
                color: {
                  argb: '00FF0000',
                  theme: 1,
                },
              },
            },
          ],
        } 
        sheet.getCell('F13').value ={
          richText: [
            {
              text: ` datum `,
              font: {
                color: {
                  argb: '00FF0000',
                  theme: 1,
                },
              },
            },
          ],
        } 
        sheet.getCell('G13').value ={
          richText: [
            {
              text:` datum `,
              font: {
                color: {
                  argb: '00FF0000',
                  theme: 1,
                },
              },
            },
          ],
        }
        sheet.getCell('H13').alignment = { wrapText: true , horizontal: 'center' }  
        sheet.getCell('H13').value = {
          richText: [
            {
              text: 'P:',
            },
            {
              text:`${total.PlaniraniSatiPredavanja}`,
              font: {
                color: {
                  argb: '00FF0000',
                  theme: 1,
                },
              },
            },
            {
              text: ' S:',
            },
            {
              text:`${total.PlaniraniSatiSeminari}`,
              font: {
                color: {
                  argb: '00FF0000',
                  theme: 1,
                },
              },
            },
            {
              text: '  V:',
            },
            {
              text:`${total.PlaniraniSatiVjezbe}`,
              font: {
                color: {
                  argb: '00FF0000',
                  theme: 1,
                },
              },
            },
          ],
        } 

        sheet.mergeCells('E15:G15')
        sheet.mergeCells('K15:M15')
        sheet.mergeCells('A15:A16')
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

  // ******** filling the main table ***********************************
  // ******************************************
  // ******************************************
  // ******************************************

        var len = (dataJSON.List1.length)  
        let start = 17
        for(let i=0; i <= len-1; i++){
            sheet.getCell("A" + start.toString()).alignment = { wrapText: true , horizontal: 'center' }  
            sheet.getCell("A" + start.toString()).value = `${i+1}`
            sheet.getCell("B" + start.toString()).value ={
              richText: [
                {
                  text:`${dataList[i].NastavnikSuradnikNaziv}`,
                  font: {
                    color: {
                      argb: '00FF0000',
                      theme: 1,
                    },
                  },
                },
              ],
            }  
            sheet.getCell("C" + start.toString()).alignment = { wrapText: true }
            sheet.getCell("C" + start.toString()).value ={
              richText: [
                {
                  text:`${dataList[i].ZvanjeNaziv}`,
                  font: {
                    color: {
                      argb: '00FF0000',
                      theme: 1,
                    },
                  },
                },
              ],
            }  
          
            sheet.getCell("D"+ start.toString()).value = {
              richText: [
                {
                  text:`${dataList[i].NazivNastavnikStatus}`,
                  font: {
                    color: {
                      argb: '00FF0000',
                      theme: 1,
                    },
                  },
                },
              ],
            }
            sheet.getCell("E"+ start.toString()).alignment = { wrapText: true , horizontal: 'center' } 
            sheet.getCell("E"+ start.toString()).value ={
              richText: [
                {
                  text:`${dataList[i].PlaniraniSatiPredavanja}`,
                  font: {
                    color: {
                      argb: '00FF0000',
                      theme: 1,
                    },
                  },
                },
              ],
            }  
            sheet.getCell("F"+ start.toString()).alignment = { wrapText: true , horizontal: 'center' }
            sheet.getCell("F"+ start.toString()).value ={
              richText: [
                {
                  text: `${dataList[i].PlaniraniSatiSeminari}`,
                  font: {
                    color: {
                      argb: '00FF0000',
                      theme: 1,
                    },
                  },
                },
              ],
            } 
            sheet.getCell("G"+ start.toString()).alignment = { wrapText: true , horizontal: 'center' }
            sheet.getCell("G"+ start.toString()).value = {
              richText: [
                {
                  text:`${parseInt(dataList[i].PlaniraniSatiVjezbe)}`,
                  font: {
                    color: {
                      argb: '00FF0000',
                      theme: 1,
                    },
                  },
                },
              ],
            }
            sheet.getCell("K"+ start.toString()).formula = "= E17 * H17 "
            sheet.getCell("L"+ start.toString()).formula = "= F17 * I17 "
            sheet.getCell("M"+ start.toString()).formula = "= G17 * J17 "
            sheet.getCell("N"+ start.toString()).formula = "= K17 * M17 "
            start++
         
            console.log('here ',dataList[i], dataList[i].PlaniraniSatiPredavanja)
        }

        sheet.mergeCells("A" + start.toString() +':'+ 'C' + start.toString()) 
        sheet.getCell("A" + start.toString()).alignment = { wrapText: true,  vertical: "middle", horizontal: "center" }
        sheet.getCell("A" + start.toString()).value ={
          richText: [
          {
            text:`UKUPNO`,
            font:{
              'bold': true
            }
          },
        ],
      } // has to be after for loop!
        
        sheet.getCell("E"+ start.toString()).value ={
          richText: [
            {
              text:`${total.PlaniraniSatiPredavanja}`,
              font: {
                color: {
                  argb: '00FF0000',
                  theme: 1,
                },
              },
            },
          ],
        } 
        sheet.getCell("F" + start.toString()).value ={
          richText: [
            {
              text:`${total.PlaniraniSatiSeminari}`,
              font: {
                color: {
                  argb: '00FF0000',
                  theme: 1,
                },
              },
            },
          ],
        }  
        sheet.getCell("G"+ start.toString()).value ={
          richText: [
            {
              text:`${total.PlaniraniSatiVjezbe}`,
              font: {
                color: {
                  argb: '00FF0000',
                  theme: 1,
                },
              },
            },
          ],
        }  
        sheet.getCell("E" + start.toString()).alignment = { wrapText: true ,  vertical: "middle", horizontal: "center" }
        sheet.getCell("F" + start.toString()).alignment = { wrapText: true ,  vertical: "middle", horizontal: "center" }
        sheet.getCell("G" + start.toString()).alignment = { wrapText: true ,  vertical: "middle", horizontal: "center" }
        sheet.getCell('K'+ (start+1).toString()).formula = "=SUM(K17:K24)"
        sheet.getCell('L'+ (start+1).toString()).formula = "=SUM(L17:L24)"
        sheet.getCell('M'+ (start+1).toString()).formula = "=SUM(M17:M24)"
        sheet.getCell('N'+ (start+1).toString()).formula = "=SUM(N17:N24)"

        sheet.mergeCells("A"+ (start+2).toString()+":"+"C"+ (start+3).toString())
        sheet.mergeCells("A"+ (start+6).toString()+":"+"C"+ (start+7).toString())
        sheet.mergeCells("J"+ (start+6).toString()+":"+"L"+ (start+7).toString())
        

        sheet.getCell("A"+ (start+2).toString()).value = {
          richText: [
            {
              text: `Prodekan za financije i upravljanje \r\n Prof. dr. sc. `,
            },
            {
              text:`${testData.prodekan1} `,
              font: {
                color: {
                  argb: '00FF0000',
                  theme: 1,
                },
              },
            },
          ],
        } 
        sheet.getCell("A"+ (start+6).toString()).value ={
          richText: [
            {
              text: `Prodekanica za nastavu i studentska pitanja ` + '\r\n \r\n Prof. dr. sc.'
            },
            {
              text: `${testData.prodekan2}\r\n `,
              font: {
                color: {
                  argb: '00FF0000',
                  theme: 1,
                },
              },
            },
          ],
        } 
        sheet.getCell("J"+ (start+6).toString()).value = {
          richText: [
            {
              text: `Dekan \r\n   Prof. dr. sc.`,
            },
            {
              text:` ${testData.dekan} `,
              font: {
                color: {
                  argb: '00FF0000',
                  theme: 1,
                },
              },
            },
          ],
        } 
        sheet.getCell("A"+ (start+2).toString()).alignment = { wrapText: true }
        sheet.getCell("A"+ (start+2).toString()).alignment = { wrapText: true }
        sheet.getCell("J"+ (start+6).toString()).alignment = { wrapText: true }
        sheet.getColumn('B').width = 21.11
        sheet.getColumn('C').width = 21.11 
        sheet.getColumn('D').width = 21.11 
        //return response.status(200).send(JSON.stringify(testData))
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